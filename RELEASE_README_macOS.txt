使用方法：

1. 解压 autoexcel-fill-macOS.zip
2. 把需要处理的 .xlsx 文件放入 autoexcel-fill/workspace 文件夹
3. 如果 macOS 提示“无法检查是否包含恶意软件”，先双击 first-run-unblock.command
4. 双击 run-autoexcel-fill.command

默认配置：

- 可直接修改 autoexcel-fill/config.ini
- target_date 留空表示使用电脑今天日期
- 日期支持 2026-06-10、0610、06-10、05/12；不写年份则默认当前年
- workbook 留空表示运行时从 workspace 中选择 Excel
- 运行时会询问使用今天日期还是手动输入日期
- 处理完成后窗口显示总结，详细过程在 autoexcel-fill/logs 文件夹中

注意：

- 不要只单独复制 autoexcel-fill 这个可执行文件，必须保留整个 autoexcel-fill 文件夹。
- 运行前请关闭正在处理的 Excel/WPS 文件。
- 如果出错，请查看 autoexcel-fill/autoexcel-fill-error.log。
