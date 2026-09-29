"""Optional local interface: streamlit run local_app.py."""
import json
from pathlib import Path
import streamlit as st
from report_generator import InfectionControlSupervisor, PERIOD_LABELS, LABELS

ROOT = Path(__file__).parent
st.set_page_config(page_title="تقارير وقاية", layout="wide")
st.image(str(ROOT / "public/wiqaya-logo.png"), width=150)
st.title("تقارير برامج وقاية")
year = st.number_input("السنة", 2020, 2100, 2026)
scope = st.selectbox("الفترة", list(PERIOD_LABELS) + ["custom"], format_func=lambda s: PERIOD_LABELS.get(s, "أرباع مختارة"))
if scope == "custom":
    quarters = st.multiselect("الأرباع", ["Q1","Q2","Q3","Q4"], default=["Q1","Q2"])
    scope = "compare:" + ",".join(quarters)
program = st.selectbox("البرنامج", ["all", *LABELS], format_func=lambda s: LABELS.get(s,"جميع البرامج"))
skip = st.checkbox("اعتبار أصفار Tracker بيانات غير مدخلة", value=True)
file = st.file_uploader("اختيار ملف Tracker", type=["xlsx"])
if file and st.button("إنشاء التقرير"):
    try:
        with st.status("جارٍ إنشاء التقرير") as status:
            report = InfectionControlSupervisor().run(file.getvalue(), json.loads((ROOT / "lib/facilities.json").read_text(encoding="utf-8")), year, (ROOT / "public/wiqaya-logo.png").read_bytes(), skip, lambda text: status.update(label=text), scope=scope, program=program)
        st.download_button("تحميل العرض النهائي", report.pptx_bytes, f"Weqaya-{year}-{scope.replace(':','-').replace(',','-')}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        st.json({"نتائج مطابقة":len(report.preview["entries"]), "غير مطابقة":report.preview["unmatched"]})
    except ValueError as error:
        st.error(str(error))
