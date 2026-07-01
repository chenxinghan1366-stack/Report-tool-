import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO
import uuid
import sqlite3
import os
import json

# ---------- 数据库 ----------
DB_PATH = os.path.join(os.path.dirname(__file__), "template_kb.db")

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    # 官方模板知识库
    c.execute('''CREATE TABLE IF NOT EXISTS templates (
        id TEXT PRIMARY KEY,
        province TEXT,
        city TEXT,
        district TEXT,
        report_type TEXT,
        template_name TEXT,
        template_version TEXT,
        source_url TEXT,
        source_authority TEXT,
        publish_date TEXT,
        required_fields TEXT,
        status TEXT
    )''')
    # 公司主数据
    c.execute('''CREATE TABLE IF NOT EXISTS companies (
        id TEXT PRIMARY KEY,
        company_name TEXT,
        province TEXT,
        city TEXT,
        district TEXT,
        tax_id TEXT
    )''')
    # 导出记录
    c.execute('''CREATE TABLE IF NOT EXISTS export_log (
        id TEXT PRIMARY KEY,
        company_id TEXT,
        template_id TEXT,
        generated_at TEXT,
        review_status TEXT,
        reviewer TEXT,
        reviewed_at TEXT,
        file_name TEXT
    )''')
    conn.commit()
    conn.close()

def load_templates(province=None, city=None, report_type=None):
    conn = sqlite3.connect(DB_PATH)
    query = "SELECT * FROM templates WHERE status='active'"
    params = []
    if province:
        query += " AND province=?"
        params.append(province)
    if city:
        query += " AND city=?"
        params.append(city)
    if report_type:
        query += " AND report_type=?"
        params.append(report_type)
    df = pd.read_sql(query, conn, params=params)
    conn.close()
    return df.to_dict('records') if not df.empty else []

def load_companies(province=None, city=None):
    conn = sqlite3.connect(DB_PATH)
    query = "SELECT * FROM companies WHERE status='active'"
    params = []
    if province:
        query += " AND province=?"
        params.append(province)
    if city:
        query += " AND city=?"
        params.append(city)
    df = pd.read_sql(query, conn, params=params)
    conn.close()
    return df.to_dict('records') if not df.empty else []

def save_export_log(record):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO export_log 
        (id, company_id, template_id, generated_at, review_status, reviewer, reviewed_at, file_name)
        VALUES (?,?,?,?,?,?,?,?)''',
        (record['id'], record['company_id'], record['template_id'],
         record['generated_at'], record['review_status'],
         record.get('reviewer',''), record.get('reviewed_at',''), record.get('file_name','')))
    conn.commit()
    conn.close()

def get_export_logs():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql("SELECT * FROM export_log ORDER BY generated_at DESC", conn)
    conn.close()
    return df.to_dict('records') if not df.empty else []

# ---------- 初始化示例数据 ----------
def init_sample_data():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # 检查是否有数据
    c.execute("SELECT COUNT(*) FROM templates")
    if c.fetchone()[0] == 0:
        # 官方模板示例
        templates = [
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
                'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额',
                'status': 'active'
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
                'required_fields': '单位名称,社保登记号,基数,单位金额,个人金额',
                'status': 'active'
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
                'required_fields': '纳税人识别号,公司名称,销售额,应纳税额',
                'status': 'active'
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
                'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额',
                'status': 'active'
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
                'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额',
                'status': 'active'
            },
        ]
        for t in templates:
            c.execute('''INSERT INTO templates 
                (id, province, city, district, report_type, template_name, template_version,
                 source_url, source_authority, publish_date, required_fields, status)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''',
                (t['id'], t['province'], t['city'], t['district'], t['report_type'],
                 t['template_name'], t['template_version'], t['source_url'],
                 t['source_authority'], t['publish_date'], t['required_fields'], t['status']))
    
    c.execute("SELECT COUNT(*) FROM companies")
    if c.fetchone()[0] == 0:
        companies = [
            {'id': 'c001', 'company_name': '上海科技发展有限公司', 'province': '上海', 'city': '上海市', 'district': '浦东新区', 'tax_id': '91310115MA1KXXXXX'},
            {'id': 'c002', 'company_name': '上海商贸有限公司', 'province': '上海', 'city': '上海市', 'district': '黄浦区', 'tax_id': '91310101MA1HXXXXX'},
            {'id': 'c003', 'company_name': '广州创新科技有限公司', 'province': '广东', 'city': '广州市', 'district': '天河区', 'tax_id': '91440106MA1YXXXXX'},
            {'id': 'c004', 'company_name': '北京智汇科技有限公司', 'province': '北京', 'city': '北京市', 'district': '海淀区', 'tax_id': '91110108MA01XXXXX'},
        ]
        for c in companies:
            c.execute('''INSERT INTO companies 
                (id, company_name, province, city, district, tax_id)
                VALUES (?,?,?,?,?,?)''',
                (c['id'], c['company_name'], c['province'], c['city'], c['district'], c['tax_id']))
    
    conn.commit()
    conn.close()

init_db()
init_sample_data()

# ---------- Streamlit 页面 ----------
st.set_page_config(page_title="官方模板匹配器", layout="wide")
st.title("📋 官方模板匹配器")
st.markdown("**选择地区 → 选择公司 → 选择报表类型 → 匹配官方模板 → 生成待复核版Excel**")

# 获取数据
all_companies = load_companies()
all_provinces = sorted(set(c['province'] for c in all_companies if c['province']))

# ===== 筛选区域 =====
col1, col2, col3 = st.columns(3)
with col1:
    province = st.selectbox("省份", [""] + all_provinces)
    cities = sorted(set(c['city'] for c in all_companies if c['province'] == province)) if province else []
    city = st.selectbox("城市", [""] + cities)
with col2:
    districts = sorted(set(c['district'] for c in all_companies if c['province'] == province and c['city'] == city)) if province and city else []
    district = st.selectbox("区县", [""] + districts)
    companies = [c for c in all_companies if c['province'] == province and c['city'] == city and c['district'] == district] if province and city and district else []
    company_names = [c['company_name'] for c in companies]
    selected_company_name = st.selectbox("公司", [""] + company_names)
with col3:
    report_type = st.selectbox("报表类型", ["", "增值税", "社保", "公积金", "企业所得税", "个人所得税"])
    # 获取选中公司
    selected_company = None
    for c in companies:
        if c['company_name'] == selected_company_name:
            selected_company = c
            break

# ===== 匹配模板 =====
if selected_company and report_type:
    st.markdown("---")
    st.subheader("🔍 匹配结果")
    
    templates = load_templates(province, city, report_type)
    if not templates:
        templates = load_templates(province, None, report_type)
    if not templates:
        templates = load_templates(None, None, report_type)
    
    if templates:
        matched = templates[0]  # 取第一个匹配的
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
            
            # 保存
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            # 记录导出日志
            export_id = str(uuid.uuid4())[:8]
            save_export_log({
                'id': export_id,
                'company_id': selected_company['id'],
                'template_id': matched['id'],
                'generated_at': datetime.now().isoformat(),
                'review_status': 'pending',
                'file_name': f"{selected_company['company_name']}_{report_type}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            })
            
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
    logs = get_export_logs()
    if logs:
        df_logs = pd.DataFrame(logs)
        st.dataframe(df_logs[['company_id', 'generated_at', 'review_status']])
    else:
        st.info("暂无导出记录")

# ===== 查看知识库 =====
with st.expander("📚 官方模板知识库"):
    all_templates = load_templates()
    if all_templates:
        st.dataframe(pd.DataFrame(all_templates))
    else:
        st.info("暂无模板")

st.caption(f"📁 数据库位置：`{DB_PATH}`")
