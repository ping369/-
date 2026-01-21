# 收银系统方案

本仓库用于收集和沉淀中西文双语收银系统的设计资料，支持电脑端与手机端数据同步。详细方案见 [docs/pos_system_design.md](docs/pos_system_design.md)。

## 资料清单
- `docs/pos_system_design.md`：整体架构与能力规划。
- `docs/generate_offline_cashier_html.py`：根据四份 Excel 库存文件生成离线可用的收银 HTML 页面，包含扫码、报表、标签打印等功能脚本。
