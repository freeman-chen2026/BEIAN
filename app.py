import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="填入模板生成备案表", layout="wide")
st.title("🛫 填入模板生成备案表")
st.markdown("首次上传模板后，每日只需上传数据文件即可。")

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

# ---------- 侧边栏：自定义固定值 ----------
st.sidebar.header("✏️ 自定义固定填入值（A、B列）")
default_supervision = st.sidebar.text_input("所属监管局（A列）", value="深圳局")
default_operator = st.sidebar.text_input("运行人标准名称（B列）", value="天成商务航空有限公司")

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

# ---------- 初始化 session_state ----------
if "template_wb" not in st.session_state:
    st.session_state.template_wb = None
if "header_row" not in st.session_state:
    st.session_state.header_row = None
if "data_start_row" not in st.session_state:
    st.session_state.data_start_row = None
if "group_rows" not in st.session_state:
    st.session_state.group_rows = 4  # 固定每组4行

# ---------- 模板管理 ----------
st.subheader("📂 模板管理")
if st.session_state.template_wb is None:
    st.info("首次使用请上传模板文件。")
    template_file = st.file_uploader("上传带格式的模板 Excel（含20条空行）", type=["xlsx"], key="template_upload")
    if template_file:
        try:
            wb = load_workbook(template_file)
            ws = wb.active
            # 查找表头行
            header_row = None
            for row in range(1, 5):
                if any("所属监管局" in str(cell.value) for cell in ws[row]):
                    header_row = row
                    break
            if header_row is None:
                st.error("模板中未找到表头行，请确认模板有“所属监管局”字段。")
                st.stop()
            data_start_row = header_row + 1
            st.session_state.template_wb = wb
            st.session_state.header_row = header_row
            st.session_state.data_start_row = data_start_row
            st.success("✅ 模板加载成功！现在可以上传数据文件。")
        except Exception as e:
            st.error(f"模板加载失败：{e}")
else:
    st.success("✅ 模板已加载（如需更换，请点击下方按钮重置）")
    if st.button("重新上传模板"):
        st.session_state.template_wb = None
        st.session_state.header_row = None
        st.session_state.data_start_row = None
        st.rerun()

# ---------- 数据上传 ----------
st.subheader("📊 数据上传")
data_file = st.file_uploader("上传航段数据导出 Excel", type=["xlsx"], key="data_upload")

if data_file and st.session_state.template_wb is not None:
    try:
        # 读取数据
        df_raw = parse_uploaded_file(data_file)
        st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")
        if len(df_raw) > 20:
            st.warning(f"数据条数（{len(df_raw)}）超过模板预设的20行，多余数据将被忽略。")

        # 从 session_state 获取模板信息
        wb = st.session_state.template_wb
        ws = wb.active
        header_row = st.session_state.header_row
        data_start_row = st.session_state.data_start_row
        group_rows = st.session_state.group_rows

        # 准备数据（最多20条）
        has_actual_depart = "实际出发" in df_raw.columns
        records = []
        for idx, row in df_raw.iterrows():
            if idx >= 20:
                break
            dt = pd.to_datetime(row["出发日期"])
            flight_date = f"{dt.year}/{dt.month}/{dt.day}"

            dep_city = str(row.get("出发城市", "")).strip()
            arr_city = str(row.get("到达城市", "")).strip()
            route = f"{dep_city}-{arr_city}" if dep_city and arr_city else f"{row['出发地']}-{row['到达地']}"
            run_type, oper_type = map_usage(row["用途"])

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
                "A": default_supervision,
                "B": default_operator,
                "C": flight_date,
                "D": run_type,
                "E": oper_type,
                "F": icao_type,
                "G": reg,
                "J": start_time,
                "K": est_landing,
                "L": is_landed,
                "M": actual_end if is_landed == "是" else "",
                "N": route,
            }
            records.append(record)

        # 写入数据（只修改指定列，O、H、I、P保留模板原有值）
        for i, rec in enumerate(records):
            row_num = data_start_row + i * group_rows
            ws[f"A{row_num}"] = rec["A"]
            ws[f"B{row_num}"] = rec["B"]
            ws[f"C{row_num}"] = rec["C"]
            ws[f"D{row_num}"] = rec["D"]
            ws[f"E{row_num}"] = rec["E"]
            ws[f"F{row_num}"] = rec["F"]
            ws[f"G{row_num}"] = rec["G"]
            ws[f"J{row_num}"] = rec["J"]
            ws[f"K{row_num}"] = rec["K"]
            ws[f"L{row_num}"] = rec["L"]
            ws[f"M{row_num}"] = rec["M"]
            ws[f"N{row_num}"] = rec["N"]
            # O 不写入，H、I、P 也不写入，保留模板原有

        # 保存
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # 预览
        preview_df = pd.DataFrame(records)
        st.subheader("📋 数据预览（按原始顺序，最多20条）")
        st.dataframe(preview_df, use_container_width=True)

        st.download_button(
            label="⬇️ 下载填入数据的模板文件",
            data=output,
            file_name="每日通航运行情况跟踪表_生成.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ 处理失败：{e}")
        st.exception(e)
else:
    if st.session_state.template_wb is None:
        st.info("👆 请先上传模板。")
    else:
        st.info("👆 请上传航段数据 Excel 文件。")
