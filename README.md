# 航班计划备案助手

每日上传航段 Excel，自动生成可粘贴到在线 Excel 的备案表格。

## 使用方式
1. 上传 `.xlsx` 文件
2. 查看整理后的预览
3. 复制 TSV 文本区内容，粘贴到在线 Excel（自动分列）
4. 也可下载 CSV 备用

## 本地运行
```bash
pip install -r requirements.txt
streamlit run app.py
