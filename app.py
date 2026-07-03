import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="每日通航运行情况跟踪表生成器", layout="wide")
st.title("🛫 每日通航运行情况跟踪表生成器")
st.markdown("上传航段计划 Excel，自动转换为跟踪表格式并导出。")

def find_header_row(df_raw):
    """查找包含关键列名的行索引（跳过空行）"""
    keywords = ["客户", "航班号", "飞机注册号", "出发日期"]
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
        if all(kw in row_str for kw in keywords):
            return idx
    return 0  # 若未找到，默认第一行

def parse_uploaded_file(uploaded_file):
    # 读取所有行（不设header）
    df_all = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    # 找到表头行
    header_row = find_header_row(df_all)
    # 重新读取，指定表头行
    df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row)
    # 去除全空行
    df = df.dropna(how="all")
    # 去除列名全为空或全是NaN的列
    df = df.dropna(axis=1, how="all")
    return df

# 映射用途 -> 运行种类 & 经营种类（可根据实际情况调整）
def map_usage(usage):
    usage = str(usage).strip()
    if "调机" in usage:
        return "调机飞行", "调机飞行"
    elif "载客" in usage:
        return "公务飞行", "私用飞行"   # 或 "包机飞行" 根据公司规则
    elif "共享租赁" in usage:
        return "公务飞行", "私用飞行"
    else:
        return "公务飞行", "私用飞行"

uploaded_file = st.file_uploader("上传 Excel 文件（航段数据导出格式）", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = parse_uploaded_file(uploaded_file)
        st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")

        # 构建目标表格
        records = []
        for _, row in df_raw.iterrows():
            # 日期格式化
            flight_date = pd.to_datetime(row["出发日期"]).strftime("%Y.%m.%d")
            # 航线：出发地-到达地
            route = f"{row['出发地']}-{row['到达地']}"
            # 用途映射
            run_type, oper_type = map_usage(row["用途"])
            # 开始时间、预计落地（时间字符串）
            start_time = str(row["计划出发"]).strip()
            est_landing = str(row["预计到达"]).strip()
            # 如果是时间对象则转换
            if hasattr(row["计划出发"], "strftime"):
                start_time = row["计划出发"].strftime("%H:%M:%S")
            if hasattr(row["预计到达"], "strftime"):
                est_landing = row["预计到达"].strftime("%H:%M:%S")

            record = {
                "深圳监管局": "深圳局",
                "运行人标准名称": "天成商务航空有限公司",
                "飞行活动的日期": flight_date,
                "当日飞行的运行种类": run_type,
                "当日飞行的经营种类": oper_type,
                "航空器型号": "",  # 可留空或从注册号推断（暂无）
                "航空器注册号": row["飞机注册号"],
                "是否向监控中心完成计划备案": "是",
                "是否获得飞行计划部门批准飞行": "是",
                "飞行开始时间": start_time,
                "飞行预计落地时间": est_landing,
                "是否已落地": "否",
                "飞行实际结束时间": "",  # 空
                "飞行地点（航线）": route,
                "选择允许的运行种类": "公务航空运行",
                "监管局是否电话跟踪该飞行动态": ""  # 可空
            }
            records.append(record)

        df_output = pd.DataFrame(records)
        # 按飞行活动日期、开始时间排序
        df_output = df_output.sort_values(["飞行活动的日期", "飞行开始时间"]).reset_index(drop=True)

        # 展示预览
        st.subheader("📋 转换后的跟踪表（预览）")
        st.dataframe(df_output, use_container_width=True)

        # 提供下载为 Excel（与第二个文件格式一致）
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

        # 同时提供 CSV 下载（备用）
        st.download_button(
            label="⬇️ 下载 CSV 文件",
            data=df_output.to_csv(index=False, encoding='utf-8-sig'),
            file_name="每日通航运行情况跟踪表.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ 处理文件时出错：{e}")
        st.info("请确认 Excel 文件包含以下列：客户、航班号、飞机注册号、用途、出发日期、计划出发、预计到达、出发地、到达地")
else:
    st.info("👆 请上传航段数据导出 Excel 文件")
