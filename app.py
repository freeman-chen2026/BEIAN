import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="每日通航运行情况跟踪表生成器", layout="wide")
st.title("🛫 每日通航运行情况跟踪表生成器")
st.markdown("上传航段计划 Excel，自动转换为跟踪表格式并导出。")

# ---------- 内置注册号 -> ICAO 机型映射（已按您的最新数据更新） ----------
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
    """查找包含关键列名的行索引（跳过空行）"""
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

# ---------- 侧边栏：映射管理 ----------
st.sidebar.header("✈️ 注册号 → ICAO 机型映射")
st.sidebar.markdown("若内置映射不全，请在下表补充（每行一个 `注册号 ICAO`，空格或Tab分隔）")
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
# 合并内置（用户输入优先）
for k, v in DEFAULT_ICAO_MAP.items():
    if k not in icao_map:
        icao_map[k] = v

# ---------- 主界面 ----------
uploaded_file = st.file_uploader("上传 Excel 文件（航段数据导出格式）", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = parse_uploaded_file(uploaded_file)
        st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")

        records = []
        for _, row in df_raw.iterrows():
            flight_date = pd.to_datetime(row["出发日期"]).strftime("%Y.%m.%d")
            # 航线 = 出发城市 - 到达城市（若城市名为空则回退四字码）
            dep_city = str(row.get("出发城市", "")).strip()
            arr_city = str(row.get("到达城市", "")).strip()
            if dep_city and arr_city:
                route = f"{dep_city}-{arr_city}"
            else:
                route = f"{row['出发地']}-{row['到达地']}"

            run_type, oper_type = map_usage(row["用途"])

            # 时间处理
            start_time = str(row["计划出发"]).strip()
            est_landing = str(row["预计到达"]).strip()
            if hasattr(row["计划出发"], "strftime"):
                start_time = row["计划出发"].strftime("%H:%M")
            if hasattr(row["预计到达"], "strftime"):
                est_landing = row["预计到达"].strftime("%H:%M")

            reg = str(row["飞机注册号"]).strip().upper()
            icao_type = icao_map.get(reg, "")  # 若未映射则为空

            record = {
                "深圳监管局": "深圳局",
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
                "是否已落地": "否",
                "飞行实际结束时间": "",
                "飞行地点（航线）": route,
                "选择允许的运行种类": "公务航空运行",
                "监管局是否电话跟踪该飞行动态": ""
            }
            records.append(record)

        df_output = pd.DataFrame(records)
        df_output = df_output.sort_values(["飞行活动的日期", "飞行开始时间"]).reset_index(drop=True)

        st.subheader("📋 转换后的跟踪表（预览）")
        st.dataframe(df_output, use_container_width=True)

        # 导出 Excel
        def to_excel_bytes(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name="Sheet1", index=False)
            return output.getvalue()

        excel_data = to_excel_bytes(df_output)
        st.download_button(
            label="⬇️ 下载 Excel 文件（.xlsx）",
            data=excel_data,
            file_name="每日通航运行情况跟踪表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="⬇️ 下载 CSV 文件",
            data=df_output.to_csv(index=False, encoding='utf-8-sig'),
            file_name="每日通航运行情况跟踪表.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ 处理文件时出错：{e}")
        st.info("请确认 Excel 包含列：客户、航班号、飞机注册号、用途、出发日期、计划出发、预计到达、出发地、到达地、出发城市、到达城市")
else:
    st.info("👆 请上传航段数据导出 Excel 文件")
