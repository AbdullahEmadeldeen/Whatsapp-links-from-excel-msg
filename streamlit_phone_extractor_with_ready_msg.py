# streamlit_phone_extractor.py
# -------------------------------------------------------------
# تطبيق Streamlit لاستخراج أرقام الموبايل من ملف .xlsx وبناء روابط واتساب
# مع رسالة عربية جاهزة + عرض الروابط كأزرار (ليس كأعمدة).
#
# مخطط الملف (سطر العناوين هو الصف الأول):
#   - العمود 1: ملاحظات (تُضاف كما هي إلى جملة الكميات)
#   - العمود 4: الإجمالي (total)
#   - الأعمدة 5..12: أسماء الأصناف (من صف العناوين) والقيم/الكميات بالصفوف التالية
#   - كل صف = أوردر واحد
#
# الناتج للتنزيل: ملف إكسل بعمودين فقط (رقم الهاتف، رابط واتساب مع ?text=)
# يتطلب: streamlit, pandas, openpyxl
# -------------------------------------------------------------

from typing import Optional, List
import re
from io import BytesIO
from urllib.parse import quote
import pandas as pd
import streamlit as st

# -----------------------------
# تهيئة الصفحة + دعم RTL
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
st.caption("حمّل ملف الإكسل وفق المخطط المحدد، وسيتم استخراج الأرقام وبناء جملة الكميات ودمجها في رسالة عربية جاهزة داخل رابط واتساب. الروابط تُعرض كأزرار.")

# -----------------------------
# ثوابت الأعمدة (فهرس يبدأ من 0)
# -----------------------------
NOTES_COL_IDX = 0       # العمود 1: ملاحظات
TOTAL_COL_IDX = 3       # العمود 4: الإجمالي
ITEMS_START_IDX = 4     # العمود 5: بداية الأعمدة الخاصة بالأصناف
ITEMS_END_IDX_INC = 11  # العمود 12 (شامل)

# -----------------------------
# القالب العربي (ثابت كما طَلبت)
# -----------------------------
AR_TEMPLATE_FIXED = (
    "السلام عليكم ورحمة الله باكلم حضرتك بخصوص اوردر زبده المصريين. "
    "اوردر حضرتك متاح للتوصيل غدا ان شاء الله من 4الى8م لو مناسب لحضرتك استاذنك اللوكيشن "
    "اوردر حضرتك : {quantity} الاجمالي {total} ,,,"
)

# -----------------------------
# دوال الموبايل والروابط
# -----------------------------
EG_MOBILE_REGEX = re.compile(r'(?:\+?20)?0?1\d{9}')  # يقبل +20/20/01

def normalize_eg_phone(text: str) -> Optional[str]:
    """إيجاد رقم موبايل مصري داخل النص وتحويله لصيغة أرقام دولية بدون + : 201XXXXXXXXX"""
    if text is None:
        return None
    s = str(text)
    m = EG_MOBILE_REGEX.search(s)
    if not m:
        return None
    digits = re.sub(r"\D", "", m.group())

    # تحويل للصيغة الدولية 201XXXXXXXXX
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
    return digits  # مثال: 2010XXXXXXXX

def extract_first_phone_from_row(row: pd.Series) -> Optional[str]:
    """يفتش كل خلايا الصف ويعيد أول رقم مصري صالح بصيغة 201XXXXXXXXX"""
    for val in row:
        digits = normalize_eg_phone(val)
        if digits:
            return digits
    return None

def encode_for_whatsapp(text: str) -> str:
    """ترميز نص الرسالة للتمرير في ?text= (يدعم العربية والأسطر الجديدة)"""
    return quote(text, safe="")

def build_wa_link(phone_digits: str, message: str) -> str:
    base = f"https://wa.me/{phone_digits}"
    if message and message.strip():
        return f"{base}?text={encode_for_whatsapp(message)}"
    return base

# -----------------------------
# بناء جملة الكميات حسب المخطط
# -----------------------------
def build_quantity_phrase(row: pd.Series, headers: List[str]) -> str:
    """
    تُكوِّن جملة الكميات بالشكل:
      "[ملاحظات إن وجدت]، صنف1: قيمة، صنف2: قيمة، ..."
    - الملاحظات (العمود 1) تُضمّن كما هي.
    - الأعمدة 5..12: إذا القيمة رقم>0 أو نص غير فارغ -> تُضاف "اسم الصنف: القيمة".
    """
    parts: List[str] = []

    # ملاحظات
    note = row.iat[NOTES_COL_IDX] if len(row) > NOTES_COL_IDX else None
    if pd.notna(note) and str(note).strip():
        parts.append(str(note).strip())

    # أصناف من 5 إلى 12 (شامل)
    start, end_inc = ITEMS_START_IDX, ITEMS_END_IDX_INC
    upper = min(end_inc, len(row) - 1)
    for idx in range(start, upper + 1):
        item_name = headers[idx] if idx < len(headers) else f"صنف {idx+1}"
        val = row.iat[idx] if idx < len(row) else None
        if pd.isna(val):
            continue
        sval = str(val).strip()
        if not sval:
            continue
        # تجاهل صفر صريح
        if sval in ("0", "0.0"):
            continue
        parts.append(f"{item_name}: {sval}")

    return "، ".join(parts)

# -----------------------------
# تصدير ملف إكسل بعمودين فقط
# -----------------------------
def to_excel_two_cols(df: pd.DataFrame, make_clickable: bool = True) -> bytes:
    out_df = df[["رقم الهاتف", "رابط واتساب"]].copy()
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
st.subheader("رفع ملف الإكسل")
file = st.file_uploader("اختر ملف .xlsx (الصف الأول يحوي عناوين الأعمدة/أسماء الأصناف)", type=["xlsx"])

if file:
    try:
        # نقرأ العناوين من الصف الأول (header=0)
        df = pd.read_excel(file, dtype=str, header=0, engine="openpyxl")
    except ImportError:
        st.error("هذا التطبيق يحتاج الحزمة 'openpyxl' لقراءة ملفات .xlsx. أضفها إلى requirements.txt ثم أعد النشر.")
        st.stop()
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        st.stop()

    st.write("معاينة أول 20 صفًا:")
    st.dataframe(df.head(20), width="stretch")

    # التحقق المبدئي من عدد الأعمدة
    if df.shape[1] < ITEMS_END_IDX_INC + 1:
        st.warning("الملف لا يحتوي على الأعمدة المطلوبة حتى العمود 12. تحقق من المخطط (ملاحظات، إجمالي في العمود 4، الأصناف من 5 إلى 12).")

    headers: List[str] = list(df.columns)

    # معالجة الصفوف → رسائل + روابط
    rows_out = []
    seen_phones = set()

    for _, row in df.iterrows():
        # 1) رقم الهاتف (يُستخرج من أي عمود)
        phone_digits = extract_first_phone_from_row(row)
        if not phone_digits:
            continue  # لا يوجد رقم صالح

        phone_display = f"+{phone_digits}"

        # 2) الإجمالي (العمود 4)
        total_val = row.iat[TOTAL_COL_IDX] if len(row) > TOTAL_COL_IDX else ""
        total_str = "" if pd.isna(total_val) else str(total_val).strip()

        # 3) جملة الكميات (ملاحظات + أصناف 5..12)
        quantity_phrase = build_quantity_phrase(row, headers)

        # 4) الرسالة العربية الثابتة
        msg = AR_TEMPLATE_FIXED.format(quantity=quantity_phrase, total=total_str)

        # 5) رابط واتساب مع ?text=
        wa_link = build_wa_link(phone_digits, msg)

        # منع التكرار لنفس الرقم
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

    # ----------------- عرض الأزرار بدل عمود الرابط -----------------
    st.divider()
    st.subheader("النتيجة")
    st.caption("ضع ✔️ في عمود (تم) أثناء العمل. الروابط تُفتح عبر أزرار لكل عميل.")

    # نعرض فقط رقم الهاتف + تم (بدون عمود الرابط)
    edited_df = st.data_editor(
        result_df[["رقم الهاتف", "تم"]],
        width="stretch",
        num_rows="fixed",
        hide_index=True,
        column_config={
            "تم": st.column_config.CheckboxColumn("تم", help="تحديد الطلب كمكتمل"),
        },
        key="editor_upload"
    )

    # أزرار واتساب لكل صف (مع معاينة الرسالة)
    st.markdown("### أزرار واتساب لكل عميل")
    for _, r in result_df.iterrows():
        with st.expander(f"📨 {r['رقم الهاتف']}", expanded=False):
            st.code(r["الرسالة (للمعاينة)"], language="text")
            st.link_button("فتح واتساب", r["رابط واتساب"], type="primary", help="فتح محادثة واتساب مع الرسالة")

    # ----------------- التنزيل: عمودان فقط -----------------
    st.divider()
    st.subheader("تحميل ملف النتائج")
    st.caption("سيتم تنزيل ملف إكسل يحتوي **عمودين فقط**: (رقم الهاتف، رابط واتساب).")

    make_clickable = st.toggle("جعل الروابط قابلة للنقر داخل Excel (HYPERLINK)", value=True, key="clickable_upload")
    excel_bytes = to_excel_two_cols(result_df, make_clickable=make_clickable)

    st.download_button(
        label="⬇️ تنزيل ملف (رقم الهاتف، رابط واتساب)",
        data=excel_bytes,
        file_name="whatsapp_orders.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("من فضلك ارفع ملف .xlsx حسب المخطط: العمود 1 ملاحظات، العمود 4 إجمالي، الأعمدة 5..12 أصناف (أسماء الأصناف في الصف الأول).")

# تذييل
st.caption("تم بإ ❤️ باستخدام Streamlit.")
