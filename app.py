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

# ---------- 主界面 ----------
st.subheader("📂 上传文件")
data_file = st.file_uploader("上传航段数据导出 Excel（如：航段数据导出 (40).xlsx）", type=["xlsx"], key="data")
template_file = st.file_uploader("上传带格式的模板 Excel（如：每日通航运行情况跟踪表（天成商务航空有限公司）20260707.xlsx）", type=["xlsx"], key="template")

if data_file and template_file:
    try:
        # 1. 读取数据
        df_raw = parse_uploaded_file(data_file)
        st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")

        # 2. 加载模板
        template_io = BytesIO(template_file.read())
        wb = load_workbook(template_io)
        ws = wb.active

        # 3. 确定表头行（假设第2行是表头，数据从第3行开始）
        header_row = 2
        data_start_row = 3

        # 4. 清空所有旧数据行（从 data_start_row 到最后）
        # 但保留可能的合并单元格？我们删除行
        # 先获取最大行
        max_row = ws.max_row
        if max_row >= data_start_row:
            ws.delete_rows(data_start_row, max_row - data_start_row + 1)

        # 5. 准备新数据（按模板列顺序）
        # 模板列顺序（从A到P）：
        col_map = {
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
        }
        # 反向映射：列名 -> 列字母
        col_letter = {v: k for k, v in col_map.items()}

        # 读取数据并生成记录
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

        # 6. 按日期和开始时间排序
        records = sorted(records, key=lambda x: (x["飞行活动的日期"], x["飞行开始时间"]))

        # 7. 将数据写入模板，并复制样式（从表头行复制样式）
        # 获取表头行所有单元格的样式（以便应用到新行）
        header_cells = {}
        for col in range(1, 17):  # A-P
            cell = ws.cell(row=header_row, column=col)
            header_cells[col] = {
                "font": copy(cell.font),
                "fill": copy(cell.fill),
                "border": copy(cell.border),
                "alignment": copy(cell.alignment),
                "number_format": cell.number_format
            }

        # 写入数据行
        for i, rec in enumerate(records):
            row_num = data_start_row + i
            for col_letter, col_name in col_map.items():
                col_idx = ws[col_letter + "1"].column  # 获取列号
                value = rec.get(col_name, "")
                cell = ws.cell(row=row_num, column=col_idx)
                cell.value = value
                # 复制表头行的样式
                style = header_cells[col_idx]
                cell.font = style["font"]
                cell.fill = style["fill"]
                cell.border = style["border"]
                cell.alignment = style["alignment"]
                cell.number_format = style["number_format"]

        # 8. 保存到内存
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.subheader("📋 预览（前5行）")
        # 显示预览：读取生成的临时文件用pandas预览，但格式可能丢失，我们直接显示DataFrame
        # 为了预览，从records构建DataFrame
        preview_df = pd.DataFrame(records)
        st.dataframe(preview_df, use_container_width=True)

        st.download_button(
            label="⬇️ 下载带格式的 Excel 文件",
            data=output,
            file_name="每日通航运行情况跟踪表_生成.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ 处理失败：{e}")
        st.info("请确保模板文件包含表头（第2行），且列标题与要求一致。")
else:
    st.info("请上传 **数据源** 和 **模板** 两个文件。")
