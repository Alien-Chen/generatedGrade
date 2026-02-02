import streamlit as st
import pandas as pd
import io
import zipfile
import os


# ==========================================
# 核心逻辑：将单个 Excel 转换为 HTML 字符串
# ==========================================
def process_excel_to_html_content(excel_file):
    """
    读取 Excel 文件对象，返回生成的 HTML 字符串内容。
    如果出错，返回 None 和 错误信息。
    """
    try:
        # 读取 Excel (注意：这里直接读取内存中的文件对象)
        # 尝试读取两个特定的 Sheet
        try:
            df_detail = pd.read_excel(excel_file, sheet_name='小题分', header=[0, 1])
            df_summary = pd.read_excel(excel_file, sheet_name='题型均分', header=[0, 1])
        except ValueError:
            return None, "找不到指定的 Sheet 名称，请确保包含 '小题分' 和 '题型均分'。"

        # --- 1. 数据预处理 ---
        def get_col_value(row, col_name_level0):
            for col in row.index:
                if str(col[0]).strip() == col_name_level0:
                    return row[col]
            return ""

        # 识别题目列
        question_cols = []
        for col in df_detail.columns:
            if '小题分' in str(col[1]):
                question_cols.append(col[0])

        # --- 2. HTML 头部 ---
        html_content = """
        <!DOCTYPE html>
        <html lang="zh-CN">
        <head>
            <meta charset="UTF-8">
            <title>学生成绩单</title>
            <style>
                body {
                    font-family: "KaiTi", "楷体", "STKaiti", serif; font-size: 14px; color: #000;
                }
                .container {
                    width: 800px; margin: 0 auto 30px auto;
                    border: 2px solid #333;
                    padding: 15px;
                    box-sizing: border-box;
                    page-break-inside: avoid;
                }
                .header-title {
                    text-align: center; font-weight: bold; font-size: 18px; margin-bottom: 15px;
                }
                .info-row {
                    display: flex; justify-content: space-between; margin-bottom: 5px; font-weight: bold;
                    padding: 0 5px;
                }
                .score-table {
                    width: 100%; border-collapse: collapse; margin-bottom: 5px; text-align: center;
                }
                .score-table td {
                    border: 1px solid #999; padding: 4px 2px; width: 8.33%;
                }
                .table-header {
                    background-color: #f2f2f2; font-weight: bold;
                }
                .summary-box {
                    text-align: center; font-weight: bold; font-size: 16px;
                    margin: 10px 0;
                    padding: 5px 0;
                    border-top: 2px dashed #ccc;
                    border-bottom: 2px dashed #ccc;
                }
                .footer {
                    display: flex; justify-content: space-between; margin-top: 20px;
                    padding: 0 20px;
                }
                .underline {
                    display: inline-block; width: 100px; border-bottom: 1px solid #000;
                }
                /* 打印分页设置 */
                @media print {
                    .page-break {
                        page-break-after: always;
                        break-after: page;
                    }
                    .container {
                        border: 2px solid #000; /* 打印时确保边框清晰 */
                    }
                }
            </style>
        </head>
        <body>
        """

        # --- 3. 循环生成每个学生的数据 ---
        for index, row in df_detail.iterrows():
            name = get_col_value(row, '姓名')
            student_id = get_col_value(row, '学号')
            class_name = get_col_value(row, '班级')
            student_id_str = str(student_id).strip()

            # 匹配汇总表
            summary_row = None
            for idx, s_row in df_summary.iterrows():
                s_id = get_col_value(s_row, '学号')
                if str(s_id).strip() == student_id_str:
                    summary_row = s_row
                    break
            
            # 获取汇总分
            if summary_row is not None:
                try:
                    def get_score(header_name):
                        for col in df_summary.columns:
                            if str(col[0]) == header_name and ('均分' in str(col[1]) or '得分' in str(col[1])):
                                val = summary_row[col]
                                # 处理"缺"字的情况
                                if pd.isna(val):
                                    return 0
                                val_str = str(val).strip()
                                if val_str == "缺":
                                    return "缺"
                                try:
                                    return float(val_str)
                                except:
                                    return 0
                        return 0
                    score_sel = get_score('选择题')
                    score_ans = get_score('解答题')
                    score_mul = get_score('多选题')
                    score_total = get_score('总分')
                except:
                    score_sel, score_ans, score_mul, score_total = 0, 0, 0, 0
            else:
                score_sel, score_ans, score_mul, score_total = 0, 0, 0, 0

            # 获取详情分
            q_data = []
            for q_label in question_cols:
                try:
                    target_col = None
                    for col in df_detail.columns:
                        if col[0] == q_label and '小题分' in str(col[1]):
                            target_col = col
                            break
                    val = row[target_col] if target_col else ""
                    if pd.isna(val):
                        val_str = ""
                    else:
                        # 处理"缺"字的情况
                        val_str = str(val).strip()
                        if val_str == "缺":
                            # 直接保留"缺"字
                            pass
                        else:
                            # 尝试转换为数值
                            try:
                                val_float = float(val_str)
                                val_str = f"{val_float:.0f}" if val_float.is_integer() else f"{val_float:.1f}"
                            except:
                                # 如果转换失败，保留原始值
                                pass
                except:
                    val_str = ""
                q_data.append({'q': str(q_label), 's': val_str})

            # 拆分两行
            split_idx = 12
            row1 = q_data[:split_idx]
            row2 = q_data[split_idx:]
            while len(row1) < 12: row1.append({'q': '&nbsp;', 's': '&nbsp;'})
            while len(row2) < 12: row2.append({'q': '&nbsp;', 's': '&nbsp;'})

            def mk_row(data, is_h):
                k = 'q' if is_h else 's'
                c = 'class="table-header"' if is_h else ''
                return f"<tr {c}>" + "".join([f"<td>{x[k]}</td>" for x in data]) + "</tr>"

            # 处理总分和各题型得分的显示
            def format_score(score):
                if score == "缺":
                    return "缺"
                try:
                    return f"{float(score):g}"
                except:
                    return "0"
            
            # 拼接单个学生的HTML
            html_content += f"""
            <div class="container">
                <div class="header-title">福建省厦门第六中学（东渡校区）语文小题分成绩条</div>
                <div class="info-row">
                    <span>姓名：{name}</span>
                    <span>学号：{student_id_str}</span>
                    <span>班级：{class_name}</span>
                    <span style="color: #666;">科目：语文</span>
                </div>
                <table class="score-table">
                    <!-- 第一行题号 -->
                    <tr class="table-header">
                        {"".join([f'<td>{item["q"]}</td>' for item in row1])}
                    </tr>
                    <!-- 第一行分数 -->
                    <tr>
                        {"".join([f'<td>{item["s"]}</td>' for item in row1])}
                    </tr>
                    <!-- 第二行题号 -->
                    <tr class="table-header">
                        {"".join([f'<td>{item["q"]}</td>' for item in row2])}
                    </tr>
                    <!-- 第二行分数 -->
                    <tr>
                        {"".join([f'<td>{item["s"]}</td>' for item in row2])}
                    </tr>
                </table>
                <div class="summary-box">
                    总分：{format_score(score_total)} (选择题{format_score(score_sel)} + 解答题{format_score(score_ans)} + 多选题{format_score(score_mul)})
                </div>
                <div class="footer">
                    <span>学生签字：<span class="underline"></span></span>
                    <span>家长签字：<span class="underline"></span></span>
                    <span>日期：<span class="underline"></span></span>
                </div>
            </div>
            <!-- 分页符 -->
            <div class="page-break"></div>
            """

        html_content += "</body></html>"
        return html_content, None

    except Exception as e:
        return None, str(e)

# ==========================================
# Streamlit 页面布局
# ==========================================
st.set_page_config(page_title="成绩条生成器", page_icon="📊", layout="centered")

st.title("📊 批量成绩条生成工具")
st.markdown("请上传包含 **'小题分'** 和 **'题型均分'** 两个Sheet的 Excel 文件。支持批量上传。")

# 文件上传控件
uploaded_files = st.file_uploader("将Excel文件拖拽到此处", type=['xlsx', 'xls'], accept_multiple_files=True)

if uploaded_files:
    # 存放处理结果的字典 {文件名: html内容}
    processed_results = {}
    
    st.write("---")
    st.write("⏳ 正在处理...")
    
    progress_bar = st.progress(0)
    
    for i, uploaded_file in enumerate(uploaded_files):
        # 1. 获取文件名 (去掉后缀)
        file_name = uploaded_file.name
        base_name = os.path.splitext(file_name)[0]
        
        # 2. 调用处理函数
        # 注意：uploaded_file 本身就是类似 open() 后的文件对象，不需要再 open
        html_content, error = process_excel_to_html_content(uploaded_file)
        
        if error:
            st.error(f"❌ 文件 **{file_name}** 处理失败: {error}")
        else:
            processed_results[f"{base_name}.html"] = html_content
            st.success(f"✅ 文件 **{file_name}** 处理成功！")
            
        # 更新进度条
        progress_bar.progress((i + 1) / len(uploaded_files))

    st.write("---")

    # 根据处理结果生成下载按钮
    if processed_results:
        # 情况 A: 只有一个文件，直接提供 HTML 下载
        if len(processed_results) == 1:
            filename, content = list(processed_results.items())[0]
            st.download_button(
                label=f"📥 下载 {filename}",
                data=content,
                file_name=filename,
                mime="text/html"
            )
        
        # 情况 B: 多个文件，打包成 Zip 下载
        else:
            st.subheader("🎉 处理完成！")
            
            # 在内存中创建 Zip 文件
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for filename, content in processed_results.items():
                    zip_file.writestr(filename, content)
            
            st.download_button(
                label="📦 批量下载所有生成的 HTML (Zip压缩包)",
                data=zip_buffer.getvalue(),
                file_name="所有成绩单.zip",
                mime="application/zip"
            )