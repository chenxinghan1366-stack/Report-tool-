def smart_column_mapping(df, required_cols, user_mapping=None):
    df_cols = list(df.columns)
    # ===== 关键：扩展同义词库，覆盖您Excel中的列名 =====
    synonyms = {
        '公司': ['公司', '企业', '单位', '公司名称', '企业名称', '单位名称', 'name', 'company', '公司名', '所属公司', '公司全称'],
        '城市': ['城市', '市', 'city', '地区', '所属城市', '所在地', '城市名'],
        '姓名': ['姓名', '名字', '员工', 'name', 'employee', '人员', '员工编号', '员工姓名', '姓名'],
        '工资基数': ['工资基数', '基数', '工资', '月工资', '基本工资', 'base', 'salary', '月应发工资', '应发工资', '月薪'],
        '社保基数': ['社保基数', '养老基数', '社保', 'social_security', '社保缴费基数'],
        '公积金基数': ['公积金基数', '公积金', 'fund', '公积金缴费基数']
    }
    
    if user_mapping:
        mapping = {k: v for k, v in user_mapping.items() if v}
        for std_name in required_cols:
            if std_name not in mapping or not mapping[std_name]:
                for df_col in df_cols:
                    if df_col in mapping.values():
                        continue
                    # 优先精确匹配
                    if df_col == std_name:
                        mapping[std_name] = df_col
                        break
                    # 同义词匹配
                    for syn in synonyms.get(std_name, []):
                        if syn.lower() in df_col.lower() or df_col.lower() in syn.lower():
                            mapping[std_name] = df_col
                            break
                    if std_name in mapping and mapping[std_name]:
                        break
        return mapping
    else:
        mapping = {}
        for std_name in required_cols:
            matched = False
            for df_col in df_cols:
                # 优先精确匹配
                if df_col == std_name:
                    mapping[std_name] = df_col
                    matched = True
                    break
                # 同义词匹配
                for syn in synonyms.get(std_name, []):
                    if syn.lower() in df_col.lower() or df_col.lower() in syn.lower():
                        mapping[std_name] = df_col
                        matched = True
                        break
                if matched:
                    break
            # 如果没匹配到，留空让用户手动选
            if std_name not in mapping:
                mapping[std_name] = None
        return mapping
