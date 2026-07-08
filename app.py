import streamlit as st
import pandas as pd
from io import StringIO

st.set_page_config(page_title="每日通航运行情况跟踪表生成器", layout="wide")
st.title("🛫 每日通航运行情况跟踪表生成器")
st.markdown("上传航段数据 Excel，生成可复制粘贴的表格文本。")

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

# ---------- 主界面 ----------
uploaded_file = st.file_uploader("上传航段数据导出 Excel", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = parse_uploaded_file(uploaded_file)
        st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")

        # 检查必要列
        required_cols = ["客户", "航班号", "飞机注册号", "用途", "出发日期", "计划出发", "预计到达",
                         "出发地", "到达地", "出发城市", "到达城市", "航段状态"]
        missing = [c for c in required_cols if c not in df_raw.columns]
        if missing:
            st.warning(f"缺少以下列（可能影响部分功能）：{missing}，请检查数据。")

        has_actual_depart = "实际出发" in df_raw.columns
        if not has_actual_depart:
            st.info("注意：Excel 中没有'实际出发'列，将使用'计划出发'作为飞行开始时间。")

        # 生成数据行（按原始顺序）
        records = []
        for _, row in df_raw.iterrows():
            dt = pd.to_datetime(row["出发日期"])
            flight_date = f"{dt.year}/{dt.month}/{dt.day}"  # 2026/7/8

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

        # ---------- 生成纯文本表格（制表符分隔） ----------
        # 定义列标题（与模板一致）
        headers = [
            "所属监管局", "运行人标准名称", "飞行活动的日期", "当日飞行的运行种类", "当日飞行的经营种类",
            "航空器型号", "航空器注册号", "是否向监控中心完成计划备案", "是否获得飞行计划部门批准飞行",
            "飞行开始时间", "飞行预计落地时间", "是否已落地", "飞行实际结束时间", "飞行地点（航线）",
            "选择允许的运行种类", "监管局是否电话跟踪该飞行动态"
        ]
        # 构建 TSV 字符串
        tsv_lines = []
        tsv_lines.append("\t".join(headers))
        for rec in records:
            row = [str(rec.get(h, "")) for h in headers]
            tsv_lines.append("\t".join(row))
        tsv_text = "\n".join(tsv_lines)

        # ---------- 展示预览（DataFrame） ----------
        df_output = pd.DataFrame(records)
        st.subheader("📋 数据预览（按原始顺序）")
        st.dataframe(df_output, use_container_width=True)

        # ---------- 可复制文本区域 ----------
        st.subheader("📄 复制以下文本，粘贴到 Excel 中（自动分列）")
        st.text_area(
            label="全选 Ctrl+A 后复制 Ctrl+C",
            value=tsv_text,
            height=300,
            key="tsv_area"
        )
        st.caption("💡 提示：复制后粘贴到 Excel，数据会自动按列分开（若未分开，可使用“数据”->“分列”功能，分隔符选“制表符”）。")

        # ---------- CSV 下载（备用） ----------
        csv_data = df_output.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="⬇️ 下载 CSV 文件（可直接用 Excel 打开）",
            data=csv_data,
            file_name="每日通航运行情况跟踪表.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ 处理失败：{e}")
        st.exception(e)
else:
    st.info("👆 请上传航段数据 Excel 文件。")
