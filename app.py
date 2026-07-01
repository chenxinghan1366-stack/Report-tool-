import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO
import uuid

# ========== 数据定义 ==========
TEMPLATES = [
    {
        'id': 't001',
        'province': '上海',
        'city': '上海市',
        'district': '浦东新区',
        'report_type': '增值税',
        'template_name': '上海市增值税纳税申报表（一般纳税人）',
        'template_version': 'v2024.1',
        'source_url': 'https://shanghai.chinatax.gov.cn/bsfw/bszn/2024/zzs.xlsx',
        'source_authority': '国家税务总局上海市税务局',
        'publish_date': '2024-01-15',
        'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额'
    },
    {
        'id': 't002',
        'province': '上海',
        'city': '上海市',
        'district': '浦东新区',
        'report_type': '社保',
        'template_name': '上海市社会保险费申报表（月度）',
        'template_version': 'v2024.1',
        'source_url': 'https://rsj.sh.gov.cn/sbjb/2024/sb.xlsx',
        'source_authority': '上海市人力资源和社会保障局',
        'publish_date': '2024-01-10',
        'required_fields': '单位名称,社保登记号,基数,单位金额,个人金额'
    },
    {
        'id': 't003',
        'province': '上海',
        'city': '上海市',
        'district': '黄浦区',
        'report_type': '增值税',
        'template_name': '上海市增值税纳税申报表（小规模纳税人）',
        'template_version': 'v2024.1',
        'source_url': 'https://shanghai.chinatax.gov.cn/bsfw/bszn/2024/zzs_small.xlsx',
        'source_authority': '国家税务总局上海市税务局',
        'publish_date': '2024-01-15',
        'required_fields': '纳税人识别号,公司名称,销售额,应纳税额'
    },
    {
        'id': 't004',
        'province': '广东',
        'city': '广州市',
        'district': '天河区',
        'report_type': '增值税',
        'template_name': '广东省增值税纳税申报表',
        'template_version': 'v2024.1',
        'source_url': 'https://guangdong.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
        'source_authority': '国家税务总局广东省税务局',
        'publish_date': '2024-01-20',
        'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额'
    },
    {
        'id': 't005',
        'province': '北京',
        'city': '北京市',
        'district': '海淀区',
        'report_type': '增值税',
        'template_name': '北京市增值税纳税申报表（一般纳税人）',
        'template_version': 'v2024.2',
        'source_url': 'https://beijing.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
        'source_authority': '国家税务总局北京市税务局',
        'publish_date': '2024-02-01',
        'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额'
    },
]

COMPANIES = [
    {'id': 'c001', 'company_name': '上海科技发展有限公司', 'province': '上海', 'city': '上海市', 'district': '浦东新区', 'tax_id': '91310115MA1KXXXXX'},
    {'id': 'c002', 'company_name': '上海商贸有限公司', 'province': '上海', 'city': '上海市', 'district': '黄浦区', 'tax_id': '91310101MA1HXXXXX'},
    {'id': 'c003', 'company_name': '广州创新科技有限公司', 'province': '广东', 'city': '广州市', 'district': '天河区', 'tax_id': '91440106MA1YXXXXX'},
    {'id': 'c004', 'company_name': '北京智汇科技有限公司', 'province': '北京', 'city': '北京市', 'district': '海淀区', 'tax_id': '91110108MA01XXXXX'},
]

# ========== 初始化 session_state ==========
if 'export_history' not in st.session_state:
    st.session_state.export_history = []

# ========== 页面 ==========
st.set_page_config(page_title="官方模板匹配器", layout="wide")
st.title("📋 官方模板匹配器")
st.markdown("**选择地区 → 选择公司 → 选择报表类型 → 匹配官方模板 → 生成待复核版Excel**")

# ===== 筛选 =====
col1, col2, col3 = st.columns(3)
with col1:
    provinces = sorted(set(c['province'] for c in COMPANIES))
    province = st.selectbox("省份", [""] + provinces)
    cities = sorted(set(c['city'] for c in COMPANIES if c['province'] == province)) if province else []
    city = st.selectbox("城市", [""] + cities)
with col2:
    districts = sorted(set(c['district'] for c in COMPANIES if c['province'] == province and c['city'] == city)) if province and city else []
    district = st.selectbox("区县", [""] + districts)
    companies = [c for c in COMPANIES if c['province'] == province and c['city'] == city and c['district'] == district]
    company_names = [c['company_name'] for c in companies]
    selected_company_name = st.selectbox("公司", [""] + company_names)
with col3:
    report_type = st.selectbox("报表类型", ["", "增值税", "社保", "公积金", "企业所得税", "个人所得税"])

# 获取选中的公司
selected_company = None
for c in companies:
    if c['company_name'] == selected_company_name:
        selected_company = c
        break

# ===== 匹配模板 =====
if selected_company and report_type:
    st.markdown("---")
    st.subheader("🔍 匹配结果")

    # 匹配逻辑：先精确匹配地区+报表类型，再降级到城市级，最后省级
    matched = None
    for t in TEMPLATES:
        if t['province'] == province and t['city'] == city and t['district'] == district and t['report_type'] == report_type:
            matched = t
            break
    if not matched:
        for t in TEMPLATES:
            if t['province'] == province and t['city'] == city and t['report_type'] == report_type:
                matched = t
                break
    if not matched:
        for t in TEMPLATES:
            if t['province'] == province and t['report_type'] == report_type:
                matched = t
                break

    if matched:
        st.success("✅ 已匹配到官方模板")
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**📄 模板信息**")
            st.write(f"模板名称：{matched['template_name']}")
            st.write(f"版本：{matched['template_version']}")
            st.write(f"发布机构：{matched['source_authority']}")
            st.write(f"发布日期：{matched['publish_date']}")
            st.write(f"必填字段：{matched['required_fields']}")
        with col_b:
            st.markdown("**🔗 来源信息**")
            st.write(f"来源URL：[{matched['source_url']}]({matched['source_url']})")
            st.write(f"适用地区：{matched['province']} {matched['city']} {matched['district']}")

        # ===== 复核复选框 =====
        st.markdown("---")
        reviewed = st.checkbox("✅ 我已人工复核确认数据无误", value=False)

        if st.button("📥 生成待复核版Excel", disabled=not reviewed):
            # 生成Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "申报表"

            # 填充数据（示例）
            headers = matched['required_fields'].split(',')
            ws.append(headers)

            sample_data = {
                '纳税人识别号': selected_company['tax_id'],
                '公司名称': selected_company['company_name'],
                '销售额': '100,000.00',
                '进项税额': '13,000.00',
                '应纳税额': '0.00',
                '单位名称': selected_company['company_name'],
                '社保登记号': 'SH123456',
                '基数': '8,000.00',
                '单位金额': '1,280.00',
                '个人金额': '640.00'
            }
            row_data = [sample_data.get(h, '') for h in headers]
            ws.append(row_data)

            # ===== 添加水印（符合要求） =====
            ws.insert_rows(1)
            ws['A1'] = '【系统生成 - 待复核版】'
            ws['A1'].font = Font(color='FF0000', bold=True, size=14)
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
            ws['A1'].alignment = Alignment(horizontal='center')

            # ===== 添加模板版本信息 =====
            ws.insert_rows(2)
            ws['A2'] = f'模板名称：{matched["template_name"]}  版本：{matched["template_version"]}'
            ws['A2'].font = Font(color='666666', size=10)
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))

            # ===== 添加来源信息 =====
            ws.insert_rows(3)
            ws['A3'] = f'来源：{matched["source_authority"]}  发布日期：{matched["publish_date"]}'
            ws['A3'].font = Font(color='666666', size=10)
            ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(headers))

            # ===== 添加审计日志Sheet =====
            audit = wb.create_sheet("审计日志")
            audit.append(['操作时间', '操作类型', '操作人', '详情'])
            audit.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 'GENERATED', '系统', f'公司:{selected_company["company_name"]}, 模板:{matched["template_name"]}'])

            # 保存到内存
            output = BytesIO()
            wb.save(output)
            output.seek(0)

            # 记录导出历史
            export_record = {
                'id': str(uuid.uuid4())[:8],
                'company': selected_company['company_name'],
                'template': matched['template_name'],
                'generated_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'status': '待复核'
            }
            st.session_state.export_history.append(export_record)

            st.success("✅ Excel已生成！")
            st.download_button(
                label="📥 下载待复核版",
                data=output,
                file_name=f"{selected_company['company_name']}_{report_type}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.info("📌 该文件为【待复核版】，正式使用前请完成人工复核流程。")
    else:
        st.warning("❌ 未匹配到官方模板，请尝试其他地区或报表类型。")
else:
    if not selected_company:
        st.info("👆 请先选择公司")
    elif not report_type:
        st.info("👆 请选择报表类型")

# ===== 历史记录 =====
with st.expander("📋 导出历史记录"):
    if st.session_state.export_history:
        st.dataframe(pd.DataFrame(st.session_state.export_history))
    else:
        st.info("暂无导出记录")

# ===== 查看知识库 =====
with st.expander("📚 官方模板知识库"):
    st.dataframe(pd.DataFrame(TEMPLATES))
