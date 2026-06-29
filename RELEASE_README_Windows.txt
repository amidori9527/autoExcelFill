使用方法：

1. 解压 AutoExcelKit-Windows.zip
2. 填表：把需要处理的 .xlsx 文件放入 AutoExcelKit\workspace 文件夹
3. 对账：把上游和后台 .xlsx 文件放入 AutoExcelKit\workspace\diffOrders 文件夹
4. 填表双击 run-autoexcel-fill.bat；对账双击 run-diff-orders.bat

默认配置：

- 可直接修改 AutoExcelKit\config.ini
- target_date 留空表示使用填表工具原有默认日期逻辑
- 日期支持 2026-06-10、0610、06-10、05/12；不写年份则默认当前年
- workbook 留空表示运行时从 workspace 中选择 Excel
- 填表运行时会询问使用默认日期还是手动输入日期
- 处理完成后窗口显示总结，详细过程在 AutoExcelKit\logs 文件夹中
- 对账工具会询问目标日期，直接回车默认当天，然后自动匹配文件名：
  上游 TranDetailReport_<id>_<日期>.xlsx，后台 收款订单_<日期>.xlsx
- 同一天匹配到多组文件时会批量对比，生成汇总 HTML，不自动复制剪贴板
- 可在 config.ini 的 [diff_orders] 中设置 auto_open_html=false，关闭自动打开 HTML
- 对账结果 HTML 文件名会带上对比组文件夹名，例如 order_diff_jz663_<时间>.html

注意：

- 不要只单独复制可执行文件，必须保留整个 AutoExcelKit 文件夹。
- 运行前请关闭正在处理的 Excel/WPS 文件。
- 如果填表出错，请查看 AutoExcelKit\autoexcel-fill-error.log。
- 如果对账出错，请查看 AutoExcelKit\diff-orders-error.log。
