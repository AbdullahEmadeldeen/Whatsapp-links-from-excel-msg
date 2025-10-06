# -------------------------------------------------------------
# تطبيق Streamlit لاستخراج أرقام الموبايل وبناء روابط واتساب برسالة عربية جاهزة
# -------------------------------------------------------------

from typing import Optional, List
import re
from io import BytesIO
from urllib.parse import quote
import pandas as pd
import streamlit as st

# -----------------------------
# تهيئة الصفحة + RTL
# -----------------------------
st.set_page_config(
    page_title="مُولِّد روابط واتساب للأوردرات",
    page_icon="🟢",
    layout="wide",
)

st.markdown("""
<style>
html, body, [class*="css"]  { direction: rtl; text-align: right; }
[data-testid="stMetricValue"] { direction:ltr; }
</style>
""", unsafe_allow_html=True)

st.title("🟢 مُولِّد روابط واتساب للأوردرات")
st.caption("حمّل ملف الإكسل، وسيتم استخراج الأرقام وبناء الكميات ودمجها في رسالة عربية جاهزة داخل رابط واتساب.")

# -----------------------------
# ثوابت الأعمدة حسب المخطط
# -----------------------------
NOTES_COL_IDX = 0      # العمود 1: ملاحظات
TOTAL_COL_IDX = 3      # العمود 4: الإجمالي
ITEMS_START_IDX = 4    # العمود 5: بداية الأصناف
ITEMS_END_IDX_INC = 11 # العمود 12 (شامل)

# -----------------------------
# قالب الرسالة الثابتة
# -----------------------------
AR_TEMPLATE_FIXED = (
    "السلام عليكم ورحمة الله باكلم حضرتك بخصوص اوردر زبده المصريين. "
    "\nاوردر حضرتك متاح للتوصيل غدا ان شاء الله من 4 الى 8م لو مناسب لحضرتك استاذنك اللوكيشن. "
    "\nاوردر حضرتك: \n{quantity}"
    "\nالاجمالي: {total}"
)

# -----------------------------
# دوال استخراج وتحويل الأرقام
# -----------------------------
EG_MOBILE_REGEX = re.compile(r'(?:\+?20)?0?1\d{9}')  # يقبل +20/20/01

def normalize_eg_phone(text: str) -> Optional[str]:
    """يحول رقم مصري لصيغة 201XXXXXXXXX"""
    if text is None:
        return None
    s = str(text)
    m = EG_MOBILE_REGEX.search(s)
    if not m:
        return None
    digits = re.sub(r"\D", "", m.group())

    # تحويل للصيغة الدولية
    if digits.startswith("0") and len(digits) == 11:
        digits = "20" + digits[1:]
    elif digits.startswith("20") and len(digits) == 12:
        pass
    elif digits.startswith("1") and len(digits) == 10:
        digits = "20" + digits
    else:
        return None

    if not (digits.startswith("201") and len(digits) == 12):
        return None
    return digits

def extract_first_phone_from_row(row: pd.Series) -> Optional[str]:
    """يستخرج أول رقم مصري صالح من الصف"""
    for val in row:
        digits = normalize_eg_phone(val)
        if digits:
            return digits
    return None

def encode_for_whatsapp(text: str) -> str:
    return quote(text, safe="")

def build_wa_link(phone_digits: str, message: str) -> str:
    base = f"https://wa.me/{phone_digits}"
    if message and message.strip():
        return f"{base}?text={encode_for_whatsapp(message)}"
    return base

# -----------------------------
# بناء جملة الكميات
# -----------------------------
def build_quantity_phrase(row: pd.Series, headers: List[str]) -> str:
    """يبني جملة الكميات والملاحظات"""
    parts: List[str] = []

    # ملاحظات
    note = row.iat[NOTES_COL_IDX] if len(row) > NOTES_COL_IDX else None
    if pd.notna(note) and str(note).strip():
        parts.append(str(note).strip())

    # الأصناف
    start, end_inc = ITEMS_START_IDX, ITEMS_END_IDX_INC
    for idx in range(start, min(end_inc, len(row)-1) + 1):
        item_name = headers[idx] if idx < len(headers) else f"صنف {idx+1}"
        val = row.iat[idx] if idx < len(row) else None
        if pd.isna(val):
            continue
        sval = str(val).strip()
        if not sval or sval in ("0", "0.0"):
            continue
        parts.append(f"\n{item_name}: {sval}")

    return "، ".join(parts)

# -----------------------------
# تصدير الملف النهائي
# -----------------------------
def to_excel_two_cols(df: pd.DataFrame, make_clickable: bool = True) -> bytes:
    out_df = df[["رقم الهاتف", "رابط واتساب", "تم"]].copy()
    if make_clickable:
        out_df["رابط واتساب"] = out_df["رابط واتساب"].apply(
            lambda url: f'=HYPERLINK("{url}", "فتح واتساب")'
        )
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="WhatsApp")
    buffer.seek(0)
    return buffer.getvalue()

# -----------------------------
# واجهة المستخدم
# -----------------------------
st.subheader("📤 رفع ملف الإكسل")
file = st.file_uploader("اختر ملف .xlsx (الصف الأول يحوي العناوين/أسماء الأصناف)", type=["xlsx"])

if file:
    try:
        df = pd.read_excel(file, dtype=str, header=0, engine="openpyxl")
    except ImportError:
        st.error("التطبيق يحتاج مكتبة openpyxl لقراءة ملفات Excel.")
        st.stop()
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        st.stop()

    st.dataframe(df, width="stretch")

    # تحذير في حالة الأعمدة غير مكتملة
    if df.shape[1] < ITEMS_END_IDX_INC + 1:
        st.warning("الملف لا يحتوي على الأعمدة المطلوبة حتى العمود 12. تحقق من المخطط.")

    headers: List[str] = list(df.columns)
    rows_out = []
    seen_phones = set()

    # معالجة البيانات
    for _, row in df.iterrows():
        phone_digits = extract_first_phone_from_row(row)
        if not phone_digits:
            continue
        phone_display = f"+{phone_digits}"

        total_val = row.iat[TOTAL_COL_IDX] if len(row) > TOTAL_COL_IDX else ""
        total_str = "" if pd.isna(total_val) else str(total_val).strip()
        quantity_phrase = build_quantity_phrase(row, headers)
        msg = AR_TEMPLATE_FIXED.format(quantity=quantity_phrase, total=total_str)
        wa_link = build_wa_link(phone_digits, msg)

        if phone_display in seen_phones:
            continue
        seen_phones.add(phone_display)

        rows_out.append({
            "رقم الهاتف": phone_display,
            "الرسالة (للمعاينة)": msg,
            "رابط واتساب": wa_link,
            "تم": False,
        })

    result_df = pd.DataFrame(rows_out, columns=["رقم الهاتف", "الرسالة (للمعاينة)", "رابط واتساب", "تم"])
    st.metric("عدد الأرقام الصالحة", len(result_df))

    # ---------------------------------------
    # واجهة الطلبات (expander لكل طلب)
    # ---------------------------------------
    st.divider()
    st.subheader("📋 قائمة الطلبات")
    st.caption("يمكنك استعراض كل طلب، الضغط على زر واتساب، أو وضع علامة (تم) بعد الإرسال.")

    done_status = {}
    if not result_df.empty and all(col in result_df.columns for col in ["رقم الهاتف", "رابط واتساب"]):
        export_df = result_df[["رقم الهاتف", "رابط واتساب"]].copy()
    else:
        export_df = pd.DataFrame(columns=["رقم الهاتف", "رابط واتساب"])

    for i, row in result_df.iterrows():
        phone = row["رقم الهاتف"]
        message = row["الرسالة (للمعاينة)"]
        link = row["رابط واتساب"]

        with st.expander(f"📞 {phone}", expanded=(i < 3)):  # أول 3 طلبات مفتوحة افتراضيًا
            cols = st.columns([3, 1])
            with cols[0]:
                st.text_area("نص الرسالة:", value=message, height=120, key=f"msg_{i}")
            with cols[1]:
                st.link_button("💬 فتح واتساب", link, type="primary")
                done_status[phone] = st.checkbox("تم ✅", key=f"done_{i}")

    # ---------------------------------------
    # التصدير إلى Excel
    # ---------------------------------------
    st.divider()
    st.subheader("📦 تنزيل ملف النتائج")

    export_df = result_df[["رقم الهاتف", "رابط واتساب"]].copy()
    export_df["تم"] = [done_status.get(row["رقم الهاتف"], False) for _, row in result_df.iterrows()]

    make_clickable = st.toggle("جعل الروابط قابلة للنقر داخل Excel (HYPERLINK)", value=True ,key="make_clickable_toggle")
    excel_bytes = to_excel_two_cols(export_df, make_clickable=make_clickable)

    st.download_button(
        label="⬇️ تنزيل ملف (رقم الهاتف، رابط واتساب، تم)",
        data=excel_bytes,
        file_name="whatsapp_orders.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("من فضلك ارفع ملف .xlsx حسب المخطط: العمود 1 ملاحظات، العمود 4 إجمالي، الأعمدة 5..12 أصناف (أسماء الأصناف في الصف الأول).")
