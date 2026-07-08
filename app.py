import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from copy import copy

st.set_page_config(page_title="每日通航运行情况跟踪表生成器", layout="wide")
st.title("🛫 每日通航运行情况跟踪表生成器")
st.markdown("上传航段计划 Excel 和格式模板，自动生成带样式的备案表格。")

# ---------- 注册号 -> ICAO 机型映射 ----------
DEFAULT_ICAO_MAP = {
    "B65AP": "GLF4",
    "B652R": "GLF4",
    "B652S": "GLF4",
    "B8105": "GLEX",
    "B8309": "GLF5",
    "B652Q": "GLF4",
    "B3926": "LJ60",
    "B8160": "GLF5",
    "B8262": "GLF4",
    "B8292": "GLF5",
}

# ---------- 辅助函数 ----------
def find_header_row(df_raw):
    keywords = ["客户", "航班号", "飞机注册号", "出发日期"]
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
        if all(kw in row_str for kw in keywords):
            return idx
    return 0

def parse_uploaded_file(uploaded_file):
    df_all = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    header_row = find_header_row(df_all)
    df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row)
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    return df

def map_usage(usage):
    usage = str(usage).strip()
    if "调机" in usage:
        return "调机飞行", "调机飞行"
    elif "载客" in usage or "共享租赁" in usage:
        return "公务飞行", "私用飞行"
    else:
        return "公务飞行", "私用飞行"

def format_time(value):
    if pd.isna(value) or value == "" or value is None:
        return ""
    if isinstance(value, str):
        if ":" in value:
            parts = value.split(":")
            if len(parts) == 2:
                return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}:00"
            elif len(parts) == 3:
                return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}:{parts[2].zfill(2)}"
        return value
    if hasattr(value, "strftime"):
        return value.strftime("%H:%M:%S")
    return str(value)

# ---------- 侧边栏：映射管理 ----------
st.sidebar.header("✈️ 注册号 → ICAO 机型映射")
user_mapping_text = st.sidebar.text_area(
    "自定义映射（覆盖内置）",
    value="\n".join([f"{k} {v}" for k, v in DEFAULT_ICAO_MAP.items()]),
    height=150
)
def parse_mapping(text):
    map_dict = {}
    for line in text.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split()
        if len(parts) >= 2:
            reg, icao = parts[0], parts[1]
            map_dict[reg.upper()] = icao.upper()
    return map_dict
icao_map = parse_mapping(user_mapping_text)
for k, v in DEFAULT_ICAO_MAP.items():
    if k not in icao_map:
        icao_map[k] = v

# ---------- session_state 缓存模板 ----------
if "template_wb" not in st.session_state:
    st.session_state.template_wb = None
if "template_ws" not in st.session_state:
    st.session_state.template_ws = None
if "template_style_cache" not in st.session_state:
    st.session_state.template_style_cache = None
if "template_header_row" not in st.session_state:
    st.session_state.template_header_row = 2  # 假设表头在第2行
if "template_loaded" not in st.session_state:
    st.session_state.template_loaded = False

# ---------- 上传模板 ----------
st.subheader("📂 上传文件")
template_file = st.file_uploader(
    "【首次使用或模板更新】上传带格式的模板 Excel",
    type=["xlsx"],
    key="template_upload"
)
if template_file:
    try:
        # 加载模板
        template_io = BytesIO(template_file.read())
        wb = load_workbook(template_io)
        ws = wb.active
        
        # 提取样式缓存（从表头行复制样式，假设第2行是表头）
        header_row = 2
        style_cache = {}
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=header_row, column=col)
            style_cache[col] = {
                "font": copy(cell.font),
                "fill": copy(cell.fill),
                "border": copy(cell.border),
                "alignment": copy(cell.alignment),
                "number_format": cell.number_format
            }
        # 保存到 session_state
        st.session_state.template_wb = wb
        st.session_state.template_ws = ws
        st.session_state.template_style_cache = style_cache
        st.session_state.template_header_row = header_row
        st.session_state.template_loaded = True
        st.success("✅ 模板已加载并缓存，现在可以上传数据文件。")
    except Exception as e:
        st.error(f"模板加载失败：{e}")

# ---------- 上传数据 ----------
data_file = st.file_uploader(
    "上传航段数据导出 Excel（如：航段数据导出 (40).xlsx）",
    type=["xlsx"],
    key="data_upload"
)

if data_file:
    if not st.session_state.template_loaded:
        st.warning("⚠️ 请先上传模板文件（首次使用必须上传）。")
        st.stop()
    
    try:
        # 读取数据
        df_raw = parse_uploaded_file(data_file)
        st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")

        # 获取缓存的模板
        wb = st.session_state.template_wb
        ws = st.session_state.template_ws
        style_cache = st.session_state.template_style_cache
        header_row = st.session_state.template_header_row

        # 清除旧数据行（从 header_row+1 开始）
        max_row = ws.max_row
        if max_row > header_row:
            ws.delete_rows(header_row + 1, max_row - header_row)

        # 准备数据
        has_actual_depart = "实际出发" in df_raw.columns
        records = []
        for _, row in df_raw.iterrows():
            # 日期格式：YYYY/M/D (无前导零)
            dt = pd.to_datetime(row["出发日期"])
            flight_date = f"{dt.year}/{dt.month}/{dt.day}"
            
            dep_city = str(row.get("出发城市", "")).strip()
            arr_city = str(row.get("到达城市", "")).strip()
            route = f"{dep_city}-{arr_city}" if dep_city and arr_city else f"{row['出发地']}-{row['到达地']}"
            run_type, oper_type = map_usage(row["用途"])
            
            # 飞行开始时间：优先实际出发，若空则用计划出发
            if has_actual_depart:
                actual_depart = format_time(row.get("实际出发", ""))
                start_time = actual_depart if actual_depart else format_time(row.get("计划出发", ""))
            else:
                start_time = format_time(row.get("计划出发", ""))
            
            est_landing = format_time(row.get("预计到达", ""))
            actual_end = format_time(row.get("实际到达", "")) if "实际到达" in df_raw.columns else ""
            status = str(row.get("航段状态", "")).strip()
            is_landed = "是" if status in ["已执飞", "已完成"] else "否"
            reg = str(row["飞机注册号"]).strip().upper()
            icao_type = icao_map.get(reg, "")

            record = {
                "所属监管局": "深圳局",
                "运行人标准名称": "天成商务航空有限公司",
                "飞行活动的日期": flight_date,
                "当日飞行的运行种类": run_type,
                "当日飞行的经营种类": oper_type,
                "航空器型号": icao_type,
                "航空器注册号": reg,
                "是否向监控中心完成计划备案": "是",
                "是否获得飞行计划部门批准飞行": "是",
                "飞行开始时间": start_time,
                "飞行预计落地时间": est_landing,
                "是否已落地": is_landed,
                "飞行实际结束时间": actual_end if is_landed == "是" else "",
                "飞行地点（航线）": route,
                "选择允许的运行种类": "3.公务航空运行",
                "监管局是否电话跟踪该飞行动态": ""
            }
            records.append(record)

        # 按日期和开始时间排序
        records = sorted(records, key=lambda x: (x["飞行活动的日期"], x["飞行开始时间"]))

        # 写入数据并复制样式
        for i, rec in enumerate(records):
            row_num = header_row + 1 + i  # 从表头下一行开始
            for col_idx, col_name in enumerate([
                "所属监管局", "运行人标准名称", "飞行活动的日期", "当日飞行的运行种类", "当日飞行的经营种类",
                "航空器型号", "航空器注册号", "是否向监控中心完成计划备案", "是否获得飞行计划部门批准飞行",
                "飞行开始时间", "飞行预计落地时间", "是否已落地", "飞行实际结束时间", "飞行地点（航线）",
                "选择允许的运行种类", "监管局是否电话跟踪该飞行动态"
            ], start=1):
                cell = ws.cell(row=row_num, column=col_idx)
                cell.value = rec[col_name]
                # 复制样式
                style = style_cache[col_idx]
                cell.font = style["font"]
                cell.fill = style["fill"]
                cell.border = style["border"]
                cell.alignment = style["alignment"]
                cell.number_format = style["number_format"]

        # 保存到内存
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # 预览（仅显示数据）
        preview_df = pd.DataFrame(records)
        st.subheader("📋 预览（前5行数据）")
        st.dataframe(preview_df.head(5), use_container_width=True)

        st.download_button(
            label="⬇️ 下载带格式的 Excel 文件",
            data=output,
            file_name="每日通航运行情况跟踪表_生成.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ 处理失败：{e}")
        st.exception(e)
else:
    if st.session_state.template_loaded:
        st.info("👆 请上传航段数据 Excel 文件。")
    else:
        st.info("👆 请先上传模板文件。")
