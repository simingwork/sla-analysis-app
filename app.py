import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date, time

from client_sla_analysis import run_analysis

st.set_page_config(page_title="客户SLA未达分析", layout="wide")

st.title("📦 客户SLA未达分析工具")
st.markdown("""
### **使用步骤：**
1. 上传 1 个或多个 Excel 文件
2. 选择 SLA should date（时间段 / 单时间点）
3. 设置 cut-off 时间
4. 点击「开始分析」
5. 下载分析结果 Excel
""")
st.caption("上传Excel → 设置 SLA should date / cut off → 生成结果Excel下载")

uploaded_files = st.file_uploader(
    "上传一个或多个Excel文件（会自动合并）",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

col1, col2 = st.columns(2)

with col1:
    mode = st.radio("SLA should date 设置方式", ["时间段", "单个时间点"], horizontal=True)

with col2:
    st.write("")

sla_range = None

if mode == "单个时间点":
    c1, c2 = st.columns(2)
    with c1:
        d = st.date_input("SLA should date（日期）", value=date.today())
    with c2:
        t = st.time_input("SLA should date（时间）", value=time(0, 0))
    sla_range = (datetime(2000, 1, 1, 0, 0, 0), datetime.combine(d, t))
else:
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        d1 = st.date_input("开始日期", value=date.today())
    with c2:
        t1 = st.time_input("开始时间", value=time(0, 0), key="t1")
    with c3:
        d2 = st.date_input("结束日期", value=date.today())
    with c4:
        t2 = st.time_input("结束时间", value=time(23, 59, 59), key="t2").replace(hour=23, minute=59, second=59)
    sla_range = (datetime.combine(d1, t1), datetime.combine(d2, t2))

st.divider()

c1, c2 = st.columns(2)
with c1:
    cut_d = st.date_input("cut_off（日期）", value=date.today(), key="cut_d")
with c2:
    cut_t = st.time_input("cut_off（时间）", value=time(11, 50), key="cut_t")
cut_off = datetime.combine(cut_d, cut_t)

run_btn = st.button("开始分析", type="primary", use_container_width=True)

if run_btn:
    if not uploaded_files:
        st.error("请先上传至少一个Excel文件。")
        st.stop()

    with st.spinner("读取并合并Excel..."):
        dfs = []
        for f in uploaded_files:
            df = pd.read_excel(f)
            dfs.append(df)
        df_all = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]

    st.success(f"已加载 {len(uploaded_files)} 个文件，合并后行数：{len(df_all):,}")

    with st.spinner("运行分析逻辑..."):
        result = run_analysis(
            df_all,
            sla_should_date=sla_range,
            cut_off=cut_off,
        )

    # result 约定返回：{"output_bytes": bytes, "filename": str, "preview": {...可选...}}
    output_bytes = result["output_bytes"]
    filename = result["filename"]

    st.success("分析完成 ✅")
    st.download_button(
        label="下载结果Excel",
        data=output_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # 可选：展示预览
    preview = result.get("preview")
    if preview:
        st.subheader("预览")
        for title, pdf in preview.items():
            st.markdown(f"**{title}**")
            st.dataframe(pdf, use_container_width=True)
