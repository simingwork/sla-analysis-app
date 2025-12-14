# 客户 SLA 未达分析工具

这是一个基于 Streamlit 的网页分析工具，用于客户 SLA 未达情况分析。
用户只需要上传 Excel 文件，在网页中设置 SLA should date 和 cut-off 时间，即可自动生成分析结果，并下载 Excel 报告。

--------------------------------------------------

功能说明

1. 支持上传一个或多个 Excel 文件（自动合并数据）
2. 支持两种 SLA should date 设置方式：
   - 时间区间
   - 单个时间点
3. 支持设置 cut-off 时间
4. 自动计算 SLA 是否达标
5. 自动生成并下载分析结果 Excel，包含：
   - 明细数据
   - 整体问题归因统计
   - 按客户问题归因（AE / FBT / CBT 等）
   - 按集配站问题归因

--------------------------------------------------

项目结构

sla-analysis-app/
│
├── app.py              Streamlit 网页入口
├── sla_analysis.py     SLA 分析核心逻辑
├── requirements.txt    Python 依赖
└── README.md           项目说明文档

--------------------------------------------------

本地运行步骤

1. 安装 Python（推荐 3.9 及以上）

2. 安装依赖
在项目根目录执行：

pip install -r requirements.txt

3. 启动网页

streamlit run app.py

4. 打开浏览器
终端会显示一个地址（一般是 http://localhost:8501），在浏览器中打开即可。

--------------------------------------------------

网页使用说明

1. 打开网页
2. 上传一个或多个 Excel 文件
3. 选择 SLA should date 设置方式：
   - 时间区间：设置开始日期 / 时间 和 结束日期 / 时间
   - 单个时间点：设置日期和时间
4. 设置 cut-off 日期和时间
5. 点击「开始分析」
6. 等待分析完成后，下载生成的 Excel 文件

--------------------------------------------------

部署到线上（生成可分享网址）

推荐使用 Streamlit Community Cloud（官方免费方案）。

步骤如下：

1. 将本项目推送到 GitHub（Public 仓库）
   仓库中需包含以下文件：
   - app.py
   - sla_analysis.py
   - requirements.txt
   - README.md

2. 打开 Streamlit Cloud
   https://streamlit.io/cloud

3. 使用 GitHub 账号登录

4. 点击「New app」

5. 配置部署信息：
   - Repository：选择你的 GitHub 仓库
   - Branch：main
   - Main file path：app.py

6. 点击 Deploy

7. 等待部署完成
   系统会生成一个网址，例如：
   https://your-app-name.streamlit.app

将该链接发送给同事即可使用，无需安装 Python。

--------------------------------------------------

常见问题

1. SLA should date 的结束时间只有到 23:59？
   Streamlit 的 time_input 只支持到分钟精度，代码中已强制将结束时间修正为 23:59:59，避免边界数据丢失。

2. 本地能跑，线上报错？
   请确认 requirements.txt 中的依赖完整，并重新部署。

3. 上传多个文件后数据不对？
   工具会自动合并所有上传的 Excel 文件，请确保表结构一致。

--------------------------------------------------

作者说明

该工具用于内部 SLA 分析场景，面向非技术同事设计，
无需本地环境即可通过网页完成分析。
