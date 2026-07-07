import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border
from openpyxl.utils import get_column_letter
from copy import copy

st.set_page_config(page_title="每日通航运行情况跟踪表生成器", layout="wide")
st.title("🛫 每日通航运行情况跟踪表生成器")
st.markdown("上传航段计划 Excel 和带格式的模板，自动生成备案表格。")

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

# ---------- 解析航段数据 ----------
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

# ---------- 初始化 session_state 存储模板 ----------
if "template_loaded" not in st.session_state:
    st.session_state.template_loaded = False
    st.session_state.template_wb = None
    st.session_state.template_ws = None
    st.session_state.template_style_cache = None

# ---------- 模板加载函数 ----------
def load_template(uploaded_file):
    """加载模板并提取样式，缓存到 session_state"""
    template_io = BytesIO(uploaded_file.read())
    wb = load_workbook(template_io)
    ws = wb.active
    # 提取表头行（假设第2行）样式，用于复制到新行
    header_row = 2
    style_cache = {}
    for col in range(1, 17):  # A-P
        cell = ws.cell(row=header_row, column=col)
        style_cache[col] = {
            "font": copy(cell.font),
            "fill": copy(cell.fill),
            "border": copy(cell.border),
            "alignment": copy(cell.alignment),
            "number_format": cell.number_format
        }
    # 保存到 session_state
    st.session_state.template_loaded = True
    st.session_state.template_wb = wb
    st.session_state.template_ws = ws
    st.session_state.template_style_cache = style_cache
    st.session_state.template_header_row = header_row
    return True

# ---------- 主界面 ----------
st.subheader("📂 上传文件")

# 如果模板尚未加载，显示上传模板的按钮
if not st.session_state.template_loaded:
    template_file = st.file_uploader("首次使用请上传带格式的模板 Excel（后续无需再传）", type=["xlsx"], key="template_upload")
    if template_file:
        try:
            load_template(template_file)
            st.success("✅ 模板已加载并缓存，现在可以上传数据文件了。")
        except Exception as e:
            st.error(f"模板加载失败：{e}")
    else:
        st.info("请先上传模板文件。")
        st.stop()
else:
    st.success("✅ 模板已加载（无需重复上传）")

# 上传数据文件
data_file = st.file_uploader("上传航段数据导出 Excel（如：航段数据导出 (40).xlsx）", type=["xlsx"], key="data_upload")

if data_file and st.session_state.template_loaded:
    try:
        # 1. 读取数据
        df_raw = parse_uploaded_file(data_file)
        st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")

        # 2. 从 session_state 获取模板工作簿和工作表
        wb = st.session_state.template_wb
        ws = st.session_state.template_ws
        style_cache = st.session_state.template_style_cache
        header_row = st.session_state.template_header_row

        # 3. 清空所有旧数据行（从第3行开始）
        max_row = ws.max_row
        if max_row >= 3:
            ws.delete_rows(3, max_row - 2)  # 删除第3行到最后

        # 4. 准备数据
        records = []
        for _, row in df_raw.iterrows():
            flight_date = pd.to_datetime(row["出发日期"]).strftime("%Y-%m-%d 00:00:00")
            dep_city = str(row.get("出发城市", "")).strip()
            arr_city = str(row.get("到达城市", "")).strip()
            if dep_city and arr_city:
                route = f"{dep_city}-{arr_city}"
            else:
                route = f"{row['出发地']}-{row['到达地']}"
            run_type, oper_type = map_usage(row["用途"])
            start_time = format_time(row.get("计划出发", ""))
            est_landing = format_time(row.get("预计到达", ""))
            actual_end = format_time(row.get("实际到达", "")) if "实际到达" in df_raw.columns else ""
            status = str(row.get("航段状态", "")).strip()
            is_landed = "是" if status in ["已执飞", "已完成"] else "否"
            reg = str(row["飞机注册号"]).strip().upper()
            icao_type = icao_map.get(reg, "")

            record = {
                "A": "深圳局",
                "B": "天成商务航空有限公司",
                "C": flight_date,
                "D": run_type,
                "E": oper_type,
                "F": icao_type,
                "G": reg,
                "H": "是",
                "I": "是",
                "J": start_time,
                "K": est_landing,
                "L": is_landed,
                "M": actual_end if is_landed == "是" else "",
                "N": route,
                "O": "3.公务航空运行",
                "P": ""
            }
            records.append(record)

        # 按日期+开始时间排序
        records = sorted(records, key=lambda x: (x["C"], x["J"]))

        # 5. 写入数据行，复制样式
        for i, rec in enumerate(records):
            row_num = 3 + i
            for col_letter in "ABCDEFGHIJKLMNOP":
                col_idx = ws[col_letter + "1"].column
                cell = ws.cell(row=row_num, column=col_idx)
                cell.value = rec[col_letter]
                # 应用样式（从表头行复制）
                style = style_cache[col_idx]
                cell.font = style["font"]
                cell.fill = style["fill"]
                cell.border = style["border"]
                cell.alignment = style["alignment"]
                cell.number_format = style["number_format"]

        # 6. 保存到内存
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # 预览
        st.subheader("📋 预览（前5行数据）")
        preview_df = pd.DataFrame(records)
        if not preview_df.empty:
            preview_df = preview_df.rename(columns={
                "A": "所属监管局",
                "B": "运行人标准名称",
                "C": "飞行活动的日期",
                "D": "当日飞行的运行种类",
                "E": "当日飞行的经营种类",
                "F": "航空器型号",
                "G": "航空器注册号",
                "H": "是否向监控中心完成计划备案",
                "I": "是否获得飞行计划部门批准飞行",
                "J": "飞行开始时间",
                "K": "飞行预计落地时间",
                "L": "是否已落地",
                "M": "飞行实际结束时间",
                "N": "飞行地点（航线）",
                "O": "选择允许的运行种类",
                "P": "监管局是否电话跟踪该飞行动态"
            })
            st.dataframe(preview_df, use_container_width=True)
        else:
            st.warning("无数据可预览")

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
    if not st.session_state.template_loaded:
        st.info("请先上传模板。")
    else:
        st.info("请上传数据文件。")
