# 每日通航运行情况跟踪表生成器

将“航段数据导出”格式的 Excel 自动转换为“每日通航运行情况跟踪表”格式，方便备案。

## 功能
- 自动识别表头（跳过空行）
- 字段映射（用途 → 运行/经营种类，航线合并等）
- 生成完全符合第二份样例格式的 Excel/CSV

## 使用
1. 上传 `航段数据导出 (xxx).xlsx`
2. 预览转换后的表格
3. 下载 Excel 文件直接使用

## 本地运行
```bash
pip install -r requirements.txt
streamlit run app.py
