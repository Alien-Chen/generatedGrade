# 批量成绩条生成工具

一个基于Streamlit的批量成绩条生成工具，支持从Excel文件生成美观的成绩条HTML文件。

## 功能特性

- 支持批量上传Excel文件
- 自动处理Excel中的"缺"字情况
- 生成美观的成绩条HTML文件
- 支持批量下载生成的文件（ZIP压缩）
- 响应式设计，打印友好

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行应用

```bash
streamlit run 批量导入版本成绩单/generatedGrade.py
```

## 部署到Streamlit Cloud

1. Fork此仓库到你的GitHub账号
2. 访问 [Streamlit Cloud](https://streamlit.io/cloud)
3. 点击 "New app"
4. 选择你的GitHub仓库
5. 设置主文件路径为 `批量导入版本成绩单/generatedGrade.py`
6. 点击 "Deploy"

## 使用说明

1. 上传包含"小题分"和"题型均分"两个Sheet的Excel文件
2. 等待应用处理完成
3. 下载生成的HTML文件
4. 在浏览器中打开HTML文件，使用打印功能打印成绩条
5. 打印时记得勾选"背景图形"选项，确保表格底色显示正确

## 注意事项

- Excel文件必须包含"小题分"和"题型均分"两个Sheet
- 确保Excel文件中的列名与代码期望的一致（学号、姓名、班级等）
- 打印时记得勾选"背景图形"选项
- 如果遇到权限错误，确保使用了正确的命令行选项启动应用