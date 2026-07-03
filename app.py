import streamlit as st
import pandas as pd
from io import StringIO

st.set_page_config(page_title="航班计划备案助手", layout="wide")
st.title("🛫 航班计划备案助手")
st.markdown("上传每日导出的航段 Excel，自动生成可粘贴的备案表格。")

uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 读取 Excel（假设第一个 sheet 为数据）
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=0)
        # 去除全空行
        df_raw = df_raw.dropna(how="all")
        st.success(f"✅ 成功读取 {len(df_raw)} 条航段记录")

        # ---------- 数据清洗与合并 ----------
        # 合并日期和时间列
        df_raw["计划出发时间"] = pd.to_datetime(
            df_raw["出发日期"].astype(str) + " " + df_raw["计划出发"].astype(str),
            errors="coerce"
        )
        df_raw["预计到达时间"] = pd.to_datetime(
            df_raw["到达日期"].astype(str) + " " + df_raw["预计到达"].astype(str),
            errors="coerce"
        )

        # 按计划出发排序
        df_raw = df_raw.sort_values("计划出发时间").reset_index(drop=True)

        # 选择需要展示的列（可根据需要调整）
        output_cols = [
            "客户", "航班号", "飞机注册号",
            "出发地", "出发城市", "到达地", "到达城市",
            "计划出发时间", "预计到达时间", "预计飞行时间"
        ]
        # 确保所有列都存在
        available_cols = [col for col in output_cols if col in df_raw.columns]
        df_display = df_raw[available_cols].copy()

        # 格式化时间显示
        df_display["计划出发时间"] = df_display["计划出发时间"].dt.strftime("%Y-%m-%d %H:%M")
        df_display["预计到达时间"] = df_display["预计到达时间"].dt.strftime("%Y-%m-%d %H:%M")

        # ---------- 展示预览 ----------
        st.subheader("📋 整理后的计划（预览）")
        st.dataframe(df_display, use_container_width=True)

        # ---------- 生成可复制文本（TSV） ----------
        tsv_data = df_display.to_csv(sep="\t", index=False)
        
        st.subheader("📄 复制以下内容粘贴到在线 Excel（自动分列）")
        st.text_area(
            label="全选并 Ctrl+C 复制",
            value=tsv_data,
            height=300,
            key="tsv_area"
        )
        st.caption("💡 提示：复制后粘贴到 Excel，数据会自动按列分开。")

        # ---------- 额外下载功能 ----------
        st.download_button(
            label="⬇️ 下载为 CSV 文件（逗号分隔）",
            data=df_display.to_csv(index=False),
            file_name="备案计划.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ 处理文件时出错：{e}")
        st.info("请确认 Excel 文件格式与示例一致（列名：客户、航班号、飞机注册号、出发日期、计划出发、到达日期、预计到达……）")
else:
    st.info("👆 请上传 Excel 文件开始")
