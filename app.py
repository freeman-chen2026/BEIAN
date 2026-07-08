import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="填入模板生成备案表", layout="wide")
st.title("🛫 填入模板生成备案表")
st.markdown("上传模板和数据，自动填入指定行，保留模板所有格式。")

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

# ---------- 上传文件 ----------
st.subheader("📂 上传文件")
template_file = st.file_uploader("上传带格式的模板 Excel（含20条空行）", type=["xlsx"], key="template")
data_file = st.file_uploader("上传航段数据导出 Excel", type=["xlsx"], key="data")

if template_file and data_file:
    try:
        # 1. 读取数据
        df_raw = parse_uploaded_file(data_file)
        st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")
        if len(df_raw) > 20:
            st.warning(f"数据条数（{len(df_raw)}）超过模板预设的20行，多余数据将被忽略。")

        # 2. 加载模板
        wb = load_workbook(template_file)
        ws = wb.active

        # 3. 确定表头行（第1行，或包含“所属监管局”的行）
        header_row = None
        for row in range(1, 5):
            if any("所属监管局" in str(cell.value) for cell in ws[row]):
                header_row = row
                break
        if header_row is None:
            st.error("模板中未找到表头行，请确认模板有“所属监管局”字段。")
            st.stop()

        # 4. 数据起始行（表头下一行）
        data_start_row = header_row + 1

        # 5. 每组占多少行（从模板推断，假设每4行一组）
        #    用户可以自定义，但我们从模板中找规律：找第一个非空行，然后找下一个非空行，计算间距
        #    简单起见，我们假设用户模板中每组占4行（如2-5,6-9,10-13...）
        group_rows = 4

        # 6. 准备新数据（按原始顺序，取前20条）
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

        # 7. 写入数据（只修改每组的第一行）
        #    组首行号 = data_start_row + i * group_rows
        for i, rec in enumerate(records):
            row_num = data_start_row + i * group_rows
            # 列顺序：A到P
            col_letters = "ABCDEFGHIJKLMNOP"
            for col_letter, value in zip(col_letters, [
                rec["A"], rec["B"], rec["C"], rec["D"], rec["E"],
                rec["F"], rec["G"], rec["H"], rec["I"], rec["J"],
                rec["K"], rec["L"], rec["M"], rec["N"], rec["O"], rec["P"]
            ]):
                cell = ws[f"{col_letter}{row_num}"]
                cell.value = value
                # 注意：我们只修改值，不修改样式，所以保留原有样式

        # 8. 如果数据条数少于模板行数，多余的行保留不动（已经是空或旧数据，可能需清空？）
        #    但模板本来就是空行，所以无需处理。

        # 9. 保存到内存
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
    st.info("👆 请同时上传模板文件和航段数据文件。")
