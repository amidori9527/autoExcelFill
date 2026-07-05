版本：AutoExcelKit 0.1.2

使用方法：

1. 解压 AutoExcelKit-macOS.zip
2. 填表：把需要处理的 .xlsx 文件放入 AutoExcelKit/workspace 文件夹
3. 对账：把上游和后台 .xlsx 文件放入 AutoExcelKit/workspace/diffOrders 文件夹
4. 如果 macOS 提示“无法检查是否包含恶意软件”，先双击 first-run-unblock.command
5. 填表双击 run-autoexcel-fill.command；对账双击 run-diff-orders.command；拉取订单 Excel 双击 run-fetch-orders.command

默认配置：

- 可直接修改 AutoExcelKit/config.ini
- target_date 留空表示使用填表工具原有默认日期逻辑
- 日期支持 2026-06-10、0610、06-10、05/12；不写年份则默认当前年
- workbook 留空表示运行时从 workspace 中选择 Excel
- 填表运行时会询问使用默认日期还是手动输入日期
- 处理完成后窗口显示总结，详细过程在 AutoExcelKit/logs 文件夹中
- 对账工具会询问目标日期，直接回车默认当天，然后自动匹配文件名：
  上游 TranDetailReport_<id>_<日期>.xlsx 或 PGW_TXNDETAIL_<id>_<日期>_....xlsx，后台 收款订单_<日期>.xlsx
- 同一天匹配到多组文件时会批量对比，自动打开汇总 HTML，不自动复制剪贴板
- 文件夹内存在 代收重复支付订单_<日期>.xlsx 时，会先做上游/TP差异，再用差异订单二次匹配重复支付订单
- 可在 config.ini 的 [diff_orders] 中设置 auto_open_html=false，关闭自动打开 HTML
- 对账结果 HTML 文件名会带上对比组文件夹名，例如 order_diff_jz663_<时间>.html
- 拉取订单 Excel 的普通参数在 config.ini 的 [fetch_orders] 中填写
- 敏感登录参数请复制 AutoExcelKit/loginConf.example.ini 为 AutoExcelKit/loginConf.ini 后填写
- 上游服务端参数填 [upstream_server_login]；TP 后台参数填 [tp_backend_login]
- run-fetch-orders.command 会先登录上游并查询报表任务，token 等会话信息会写回 [upstream_server_session]

注意：

- 不要只单独复制可执行文件，必须保留整个 AutoExcelKit 文件夹。
- 运行前请关闭正在处理的 Excel/WPS 文件。
- 如果填表出错，请查看 AutoExcelKit/autoexcel-fill-error.log。
- 如果对账出错，请查看 AutoExcelKit/diff-orders-error.log。
