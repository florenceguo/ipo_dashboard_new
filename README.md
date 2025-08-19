# IPO数据看板

一个展示IPO市场数据的可视化网站，包含：

## 功能特性
- 📊 各板块月度涨跌幅趋势
- 📈 新股发行统计（数量与募资金额）
- 🎯 主板A类与B类中签率对比  
- 💰 IPO打新收益测算工具
- 📱 响应式设计，支持移动端

## 技术栈
- HTML5 + CSS3 + JavaScript
- Chart.js 图表库
- 数据来源：Wind API

## 部署
本项目已配置 Vercel 自动部署，推送到 GitHub 后会自动更新网站。

## 数据更新
数据通过 Wind API 获取，定期更新 `excel_data_optimized.json` 文件即可刷新网站数据。
