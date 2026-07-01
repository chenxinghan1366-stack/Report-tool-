import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from io import BytesIO

# ---------- 模拟数据 ----------
SOURCES = [
    {"省份": "上海", "城市": "上海市", "机构": "上海市税务局", "来源": "https://shanghai.chinatax.gov.cn/"},
    {"省份": "广东", "城市": "广州市", "机构": "广东省税务局", "来源": "https://guangdong.chinatax.gov.cn/"},
]

TEMPLATES = [
    {"省份": "上海", "城市": "上海市", "模板名称": "上海市社保申报表（月度）", "报表类型": "月度申报", "版本": "1.0", "必填字段": ["单位名称","基数"]},
    {"省份": "广东", "城市": "广州市", "模板名称": "广东省社保申报表", "报表类型": "月度申报", "版本": "1.0", "必填字段": ["单位名称","基数"]},
]

RULES = [
    {"省份": "上海", "城市": "上海市", "报表类型": "月度申报", "规则": "养老16%/8%", "来源引用": "沪人社规〔2024〕22号"},
    {"省份": "广东", "城市": "广州市", "报表类型": "月度申报", "规则": "养老14%/8%", "来源引用": "粤人社规〔2024〕8号"},
]

COMPANIES = [
    {"公司名称": "上海科技公司", "省份": "上海", "城市": "上海市", "区县": "浦东新区"},
    {"公司名称": "广州科技公司", "省份": "广东", "城市": "广州市", "区县": "天河区"},
]

st.set_page_config(page_title="报表匹配工具", layout="centered")
st.title("社保报表匹配工具")
st.markdown("---")

province = st.selectbox("省份", sorted(set(c["省份"] for c in COMPANIES)))
cities = sorted(set(c["城市"] for c in COMPANIES if c["省份"] == province))
city = st.selectbox("城市", cities)
districts = sorted(set(c["区县"] for c in COMPANIES if c["省份"] == province and c["城市"] == city))
district = st.selectbox("区县", districts)
companies = [c["公司名称"] for c in COMPANIES if c["省份"] == province and c["城市"] == city and c["区县"] == district]
company = st.selectbox("公司", companies)
report_type = st.selectbox("报表类型", ["月度申报", "年度汇算"])
year = st.selectbox("年份", [2024, 2025], index=0)
month = None
if report_type == "月度申报":
    month = st.selectbox("月份", list(range(1,13)), index=11)

if st.button("匹配官方模板"):
    template = next((t for t in TEMPLATES if t["城市"] == city and t["报表类型"] == report_type), None)
    rule = next((r for r in RULES if r["城市"] == city and r["报表类型"] == report_type), None)
    if template and rule:
        st.success("匹配成功！")
        st.write(f"模板：{template['模板名称']} (v{template['版本']})")
        st.write(f"规则：{rule['规则']}")
        st.write(f"来源：{rule['来源引用']}")
        
        if st.button("生成假数据Excel（待复核版）"):
            wb = Workbook()
            ws = wb.active
            ws.append(['公司', '险种', '基数', '单位金额', '个人金额'])
            ws.append([company, '养老保险', 8000, 1280, 640])
            ws.insert_rows(1)
            ws['A1'] = '待复核版'
            ws.merge_cells('A1:E1')
            audit = wb.create_sheet('AuditTrail')
            audit.append(['时间', '操作'])
            audit.append([datetime.now().isoformat(), '生成'])
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            st.download_button("下载Excel", data=output, file_name=f"{company}_待复核.xlsx")
    else:
        st.error("未匹配到模板")

with st.expander("查看所有数据"):
    st.dataframe(pd.DataFrame(SOURCES))
    st.dataframe(pd.DataFrame(TEMPLATES))
    st.dataframe(pd.DataFrame(RULES))
