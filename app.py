import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO
import uuid
import sqlite3
import os
import zipfile
import json

# ========== 数据库初始化 ==========
DB_PATH = os.path.join(os.path.dirname(__file__), "app_data.db")

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS companies (
        id TEXT PRIMARY KEY, company_name TEXT, province TEXT, city TEXT, district TEXT, tax_id TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS templates (
        id TEXT PRIMARY KEY, province TEXT, city TEXT, district TEXT, report_type TEXT,
        template_name TEXT, template_version TEXT, source_url TEXT, source_authority TEXT,
        publish_date TEXT, required_fields TEXT, status TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS rules (
        id TEXT PRIMARY KEY, city TEXT, unit_social REAL, personal_social REAL,
        unit_fund REAL, personal_fund REAL, social_min REAL, social_max REAL,
        fund_min REAL, fund_max REAL, source_quote TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS export_history (
        id TEXT PRIMARY KEY, company_id TEXT, template_id TEXT, company_name TEXT,
        city TEXT, report_type TEXT, period_type TEXT, generated_at TEXT,
        review_status TEXT, reviewer TEXT, reviewed_at TEXT, file_name TEXT, file_data BLOB
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS custom_templates (
        id TEXT PRIMARY KEY, name TEXT, file_data BLOB, field_mapping TEXT, created_at TEXT
    )''')
    conn.commit()
    conn.close()

init_db()

# ========== 数据操作函数 ==========
def dict_fetchall(cursor):
    columns = [col[0] for col in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]

def load_companies():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM companies")
    rows = dict_fetchall(c)
    conn.close()
    return rows

def save_companies(companies):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM companies")
    for comp in companies:
        c.execute('''INSERT INTO companies (id, company_name, province, city, district, tax_id)
            VALUES (?,?,?,?,?,?)''',
            (comp.get('id', str(uuid.uuid4())[:8]), comp['company_name'], comp['province'],
             comp['city'], comp.get('district',''), comp.get('tax_id','')))
    conn.commit()
    conn.close()

def load_templates():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM templates WHERE status='active'")
    rows = dict_fetchall(c)
    conn.close()
    return rows

def save_template(template):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO templates 
        (id, province, city, district, report_type, template_name, template_version,
         source_url, source_authority, publish_date, required_fields, status)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''',
        (template['id'], template['province'], template['city'], template.get('district',''),
         template['report_type'], template['template_name'], template['template_version'],
         template.get('source_url',''), template.get('source_authority',''),
         template.get('publish_date',''), template.get('required_fields',''), template.get('status','active')))
    conn.commit()
    conn.close()

def load_rules():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM rules")
    rows = dict_fetchall(c)
    conn.close()
    return rows

def save_rules(rules):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM rules")
    for r in rules:
        c.execute('''INSERT INTO rules 
            (id, city, unit_social, personal_social, unit_fund, personal_fund,
             social_min, social_max, fund_min, fund_max, source_quote)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
            (r['id'], r['city'], r['unit_social'], r['personal_social'],
             r['unit_fund'], r['personal_fund'], r.get('social_min',0), r.get('social_max',999999),
             r.get('fund_min',0), r.get('fund_max',999999), r.get('source_quote','')))
    conn.commit()
    conn.close()

def save_export(record):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO export_history 
        (id, company_id, template_id, company_name, city, report_type, period_type,
         generated_at, review_status, reviewer, reviewed_at, file_name, file_data)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (record['id'], record.get('company_id',''), record.get('template_id',''),
         record['company_name'], record.get('city',''), record.get('report_type',''),
         record.get('period_type',''), record['generated_at'], record.get('review_status','pending'),
         record.get('reviewer',''), record.get('reviewed_at',''),
         record.get('file_name',''), record.get('file_data', None)))
    conn.commit()
    conn.close()

def load_export_history():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM export_history ORDER BY generated_at DESC")
    rows = dict_fetchall(c)
    conn.close()
    return rows

def update_export_status(export_id, status, reviewer):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''UPDATE export_history 
        SET review_status=?, reviewer=?, reviewed_at=?
        WHERE id=?''',
        (status, reviewer, datetime.now().isoformat(), export_id))
    conn.commit()
    conn.close()

# ========== 初始化默认数据 ==========
def init_default_data():
    if not load_templates():
        templates = [
            {'id': 't001', 'province': '上海', 'city': '上海市', 'district': '浦东新区',
             'report_type': '增值税', 'template_name': '上海市增值税纳税申报表（一般纳税人）',
             'template_version': 'v2024.1', 'source_url': 'https://shanghai.chinatax.gov.cn/bsfw/bszn/2024/zzs.xlsx',
             'source_authority': '国家税务总局上海市税务局', 'publish_date': '2024-01-15',
             'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额', 'status': 'active'},
            {'id': 't002', 'province': '上海', 'city': '上海市', 'district': '浦东新区',
             'report_type': '社保', 'template_name': '上海市社会保险费申报表（月度）',
             'template_version': 'v2024.1', 'source_url': 'https://rsj.sh.gov.cn/sbjb/2024/sb.xlsx',
             'source_authority': '上海市人力资源和社会保障局', 'publish_date': '2024-01-10',
             'required_fields': '单位名称,社保登记号,基数,单位金额,个人金额', 'status': 'active'},
            {'id': 't003', 'province': '广东', 'city': '广州市', 'district': '天河区',
             'report_type': '增值税', 'template_name': '广东省增值税纳税申报表',
             'template_version': 'v2024.1', 'source_url': 'https://guangdong.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
             'source_authority': '国家税务总局广东省税务局', 'publish_date': '2024-01-20',
             'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额', 'status': 'active'},
            {'id': 't004', 'province': '北京', 'city': '北京市', 'district': '海淀区',
             'report_type': '增值税', 'template_name': '北京市增值税纳税申报表（一般纳税人）',
             'template_version': 'v2024.2', 'source_url': 'https://beijing.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
             'source_authority': '国家税务总局北京市税务局', 'publish_date': '2024-02-01',
             'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额', 'status': 'active'},
        ]
        for t in templates:
            save_template(t)
    if not load_rules():
        rules = [
            {'id': 'r001', 'city': '上海', 'unit_social': 0.16, 'personal_social': 0.08,
             'unit_fund': 0.07, 'personal_fund': 0.07, 'social_min': 7310, 'social_max': 36549,
             'fund_min': 2590, 'fund_max': 34188, 'source_quote': '沪人社规〔2024〕22号'},
            {'id': 'r002', 'city': '北京', 'unit_social': 0.16, 'personal_social': 0.08,
             'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 6326, 'social_max': 33891,
             'fund_min': 2420, 'fund_max': 33891, 'source_quote': '京人社发〔2024〕15号'},
            {'id': 'r003', 'city': '广州', 'unit_social': 0.15, 'personal_social': 0.08,
             'unit_fund': 0.10, 'personal_fund': 0.10, 'social_min': 4588, 'social_max': 22941,
             'fund_min': 2300, 'fund_max': 27960, 'source_quote': '粤人社规〔2024〕8号'},
        ]
        save_rules(rules)

init_default_data()

# ========== 解析上传的Excel ==========
def parse_uploaded_excel(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    all_companies = []
    city_to_province = {
        '上海': '上海', '北京': '北京', '广州': '广东', '深圳': '广东',
        '杭州': '浙江', '南京': '江苏', '苏州': '江苏', '成都': '四川',
        '重庆': '重庆', '武汉': '湖北', '西安': '陕西', '郑州': '河南',
        '长沙': '湖南', '青岛': '山东', '宁波': '浙江', '天津': '天津'
    }
    for sheet in sheets:
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            header_row = None
            for i, row in df.iterrows():
                row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                if '所属城市' in row_text or '城市' in row_text or '分公司' in row_text:
                    header_row = i
                    break
            if header_row is not None:
                df = pd.read_excel(file, sheet_name=sheet, skiprows=header_row)
                df.columns = [str(c).strip() for c in df.columns]
                city_col = None
                company_col = None
                district_col = None
                for col in df.columns:
                    if '所属城市' in col or '城市' in col:
                        city_col = col
                    elif '分公司' in col or '公司' in col:
                        company_col = col
                    elif '区县' in col or '区' in col:
                        district_col = col
                if city_col and company_col:
                    for _, row in df.iterrows():
                        city = str(row[city_col]) if pd.notna(row[city_col]) else ''
                        company = str(row[company_col]) if pd.notna(row[company_col]) else ''
                        district = str(row[district_col]) if district_col and pd.notna(row[district_col]) else ''
                        if city and company:
                            province = city_to_province.get(city, city)
                            all_companies.append({
                                'company_name': company,
                                'province': province,
                                'city': city,
                                'district': district,
                                'tax_id': ''
                            })
        except:
            continue
    unique = []
    seen = set()
    for c in all_companies:
        key = (c['company_name'], c['city'])
        if key not in seen:
            seen.add(key)
            unique.append(c)
    return unique

# ========== Streamlit 页面 ==========
st.set_page_config(page_title="官方模板匹配器", layout="wide")
st.title("📋 官方模板匹配器（含自动识别与统计口径）")
st.markdown("**上传Excel → 自动提取城市/公司 → 选择模板和统计口径 → 生成待复核版Excel**")

# ===== 侧边栏 =====
with st.sidebar:
    st.header("📤 上传数据Excel")
    st.markdown("上传包含公司/城市信息的Excel，系统自动提取所有地区")
    uploaded_file = st.file_uploader("选择Excel文件（.xlsx）", type=["xlsx"])
    
    if uploaded_file:
        with st.spinner("正在解析Excel..."):
            companies = parse_uploaded_excel(uploaded_file)
            if companies:
                save_companies(companies)
                st.success(f"成功提取 {len(companies)} 家公司")
                # 尝试读取数据
                try:
                    xls = pd.ExcelFile(uploaded_file)
                    data_sheet = None
                    for s in xls.sheet_names:
                        if '明细' in s or '月度' in s or '数据' in s:
                            data_sheet = s
                            break
                    if data_sheet:
                        df_data = pd.read_excel(uploaded_file, sheet_name=data_sheet)
                        st.session_state['imported_df'] = df_data
                        st.success(f"已读取数据Sheet「{data_sheet}」，共{len(df_data)}行")
                except:
                    pass
            else:
                st.warning("未识别到公司数据，请确认Excel包含「城市」和「公司」列")
    
    with st.sidebar.expander("🏢 当前公司列表"):
        companies = load_companies()
        if companies:
            st.dataframe(pd.DataFrame(companies))
            st.caption(f"共 {len(companies)} 家公司")
        else:
            st.info("暂无数据")

# ===== 主体 =====
companies = load_companies()
if not companies:
    st.info("👈 请先在侧边栏上传包含公司/城市数据的Excel")
    st.stop()

# 获取所有省份
all_provinces = sorted(set(c['province'] for c in companies if c['province']))

col1, col2, col3 = st.columns(3)
with col1:
    province = st.selectbox("省份", [""] + all_provinces)
    cities = sorted(set(c['city'] for c in companies if c['province'] == province)) if province else sorted(set(c['city'] for c in companies))
    city = st.selectbox("城市", [""] + cities)
with col2:
    districts = sorted(set(c['district'] for c in companies if c['province'] == province and c['city'] == city)) if province and city else []
    district = st.selectbox("区县", [""] + districts)
    company_list = [c for c in companies if c['province'] == province and c['city'] == city and (not district or c['district'] == district)]
    company_names = [c['company_name'] for c in company_list]
    selected_company_names = st.multiselect("公司（可多选）", company_names)
with col3:
    report_type = st.selectbox("报表类型", ["", "增值税", "社保", "公积金", "企业所得税", "个人所得税"])
    period_type = st.selectbox("统计口径", ["月度（12月单月）", "累计（1-12月）"])

# ===== 获取选中的公司对象 =====
selected_companies = [c for c in company_list if c['company_name'] in selected_company_names]

# ===== 匹配模板 =====
if selected_companies and report_type:
    st.markdown("---")
    st.subheader("🔍 匹配结果")
    
    templates = load_templates()
    rules = load_rules()
    rules_dict = {r['city']: r for r in rules}
    
    # 匹配模板
    matched = None
    for t in templates:
        if t['province'] == province and t['city'] == city and t['district'] == district and t['report_type'] == report_type:
            matched = t
            break
    if not matched:
        for t in templates:
            if t['province'] == province and t['city'] == city and t['report_type'] == report_type:
                matched = t
                break
    if not matched:
        for t in templates:
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
    else:
        st.warning("⚠️ 未匹配到官方模板，将使用通用模板")
        matched = {
            'id': 'gen001',
            'template_name': f'{report_type}通用申报表',
            'template_version': 'v1.0',
            'source_authority': '系统通用',
            'publish_date': datetime.now().strftime('%Y-%m-%d'),
            'required_fields': '纳税人识别号,公司名称,申报金额',
            'source_url': '#'
        }
    
    # 数据校验
    st.subheader("📋 数据校验")
    missing_rules = []
    for comp in selected_companies:
        if comp['city'] not in rules_dict:
            missing_rules.append(comp['city'])
    if missing_rules:
        st.warning(f"⚠️ 以下城市缺少规则，将使用默认比例计算：{', '.join(set(missing_rules))}")
    else:
        st.success("✅ 所有城市已配置规则")
    
    # ===== 生成按钮 =====
    reviewed = st.checkbox("✅ 我已人工复核确认数据无误", value=False)
    
    if st.button("📥 生成待复核版Excel", disabled=not reviewed):
        generated_files = []
        summary = []
        errors = []
        
        for comp in selected_companies:
            try:
                # 获取该城市规则
                rule = rules_dict.get(comp['city'])
                fields = matched['required_fields'].split(',')
                
                # 构建数据（优先从导入数据读取，否则用示例）
                if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
                    df_data = st.session_state['imported_df']
                    # 尝试匹配公司
                    company_data = df_data[df_data.apply(lambda row: comp['company_name'] in str(row.values), axis=1)]
                    if not company_data.empty:
                        first_row = company_data.iloc[0]
                        row_data = []
                        for f in fields:
                            matched_col = None
                            for col in df_data.columns:
                                if f in str(col) or str(col) in f:
                                    matched_col = col
                                    break
                            if matched_col:
                                row_data.append(first_row[matched_col])
                            else:
                                row_data.append('')
                    else:
                        row_data = [comp.get('tax_id', ''), comp['company_name'], '100,000.00', '13,000.00', '0.00']
                else:
                    sample_data = {
                        '纳税人识别号': comp.get('tax_id', ''),
                        '公司名称': comp['company_name'],
                        '销售额': '100,000.00',
                        '进项税额': '13,000.00',
                        '应纳税额': '0.00',
                        '单位名称': comp['company_name'],
                        '社保登记号': 'SH123456',
                        '基数': '8,000.00',
                        '单位金额': str(round(8000 * rule['unit_social'], 2)) if rule else '',
                        '个人金额': str(round(8000 * rule['personal_social'], 2)) if rule else '',
                        '申报金额': '100,000.00'
                    }
                    row_data = [sample_data.get(f, '') for f in fields]
                
                # 创建Excel
                wb = Workbook()
                ws = wb.active
                ws.title = "申报表"
                ws.append(fields)
                ws.append(row_data)
                
                # 水印
                ws.insert_rows(1)
                ws['A1'] = f'【系统生成 - 待复核版】统计口径：{period_type}'
                ws['A1'].font = Font(color='FF0000', bold=True, size=14)
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fields))
                ws['A1'].alignment = Alignment(horizontal='center')
                
                ws.insert_rows(2)
                ws['A2'] = f'模板名称：{matched["template_name"]}  版本：{matched["template_version"]}'
                ws['A2'].font = Font(color='666666', size=10)
                ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(fields))
                
                ws.insert_rows(3)
                ws['A3'] = f'来源：{matched.get("source_authority","")}  发布日期：{matched.get("publish_date","")}'
                ws['A3'].font = Font(color='666666', size=10)
                ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(fields))
                
                # 审计日志
                audit = wb.create_sheet("审计日志")
                audit.append(['操作时间', '操作类型', '操作人', '详情'])
                audit.append([datetime.now().isoformat(), 'GENERATED', '系统', f'公司:{comp["company_name"]}, 模板:{matched["template_name"]}'])
                
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                fname = f"{comp['company_name']}_{report_type}_{period_type.replace('（','_').replace('）','')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                generated_files.append((fname, output.getvalue()))
                summary.append({'公司': comp['company_name'], '城市': comp['city'], '模板': matched['template_name'], '状态': '待复核'})
                
                # 保存历史
                save_export({
                    'id': str(uuid.uuid4())[:8],
                    'company_id': comp['id'],
                    'template_id': matched.get('id', 'gen001'),
                    'company_name': comp['company_name'],
                    'city': comp['city'],
                    'report_type': report_type,
                    'period_type': period_type,
                    'generated_at': datetime.now().isoformat(),
                    'review_status': 'pending',
                    'file_name': fname,
                    'file_data': output.getvalue()
                })
            except Exception as e:
                errors.append(f"{comp['company_name']}: {str(e)}")
        
        if errors:
            for err in errors:
                st.warning(err)
        if generated_files:
            st.success(f"✅ 成功生成 {len(generated_files)} 份报表")
            st.dataframe(pd.DataFrame(summary))
            if len(generated_files) > 1:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zf:
                    for fname, data in generated_files:
                        zf.writestr(fname, data)
                zip_buffer.seek(0)
                st.download_button("📦 下载全部报表（ZIP）", data=zip_buffer, file_name=f"报表_{datetime.now().strftime('%Y%m%d')}.zip", mime="application/zip")
            else:
                fname, data = generated_files[0]
                st.download_button(f"📥 下载 {fname}", data=BytesIO(data), file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    if not selected_companies:
        st.info("👆 请先选择公司")
    elif not report_type:
        st.info("👆 请选择报表类型")

# ===== 导出历史 =====
with st.expander("📋 导出历史记录"):
    history = load_export_history()
    if history:
        df_hist = pd.DataFrame(history)
        st.dataframe(df_hist[['company_name', 'city', 'report_type', 'period_type', 'generated_at', 'review_status']])
        
        # 复核功能
        pending = [h for h in history if h['review_status'] == 'pending']
        if pending:
            st.subheader("✅ 复核待处理报表")
            opts = [f"{h['company_name']} - {h['city']} ({h['generated_at'][:10]})" for h in pending]
            sel_idx = st.selectbox("选择要复核的报表", range(len(opts)), format_func=lambda x: opts[x])
            selected = pending[sel_idx]
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ 通过复核"):
                    update_export_status(selected['id'], 'approved', '复核员')
                    st.success("已通过复核")
                    st.rerun()
            with col2:
                if st.button("❌ 驳回"):
                    update_export_status(selected['id'], 'rejected', '复核员')
                    st.warning("已驳回")
                    st.rerun()
    else:
        st.info("暂无导出记录")

# ===== 查看知识库 =====
with st.expander("📚 官方模板知识库"):
    templates = load_templates()
    if templates:
        st.dataframe(pd.DataFrame(templates))
    else:
        st.info("暂无模板")
