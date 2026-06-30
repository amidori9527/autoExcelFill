# Fetch Orders 操作文档

本文档说明如何使用 `fetch-orders` 拉取上游 `TranDetailReport` Excel 文件，并放入 `workspace/diffOrders` 供订单对账工具继续使用。

## 一、执行前准备

### 1. 准备登录配置文件

项目根目录里提供了模板文件：

```text
loginConf.example.ini
```

第一次使用前，需要复制一份并改名为：

```text
loginConf.ini
```

开发环境可以使用命令：

```bash
cp loginConf.example.ini loginConf.ini
```

打包后的 `dist/AutoExcelKit` 目录中也一样，需要把 `loginConf.example.ini` 复制成同目录下的 `loginConf.ini`。

### 2. 填写登录信息

打开 `loginConf.ini`，在 `[upstream_server_login]` 下填写以下四项：

```ini
[upstream_server_login]
api_key = 已经默认填好，不要删除
username =
password =
institution_id =
```

其中 `api_key` 已经在模板里默认填好，除非维护人员要求，否则不要修改。用户只需要填写下面三项：

- `username`：上游后台登录用户名
- `password`：上游后台登录密码
- `institution_id`：上游后台登录所需机构 ID

`ip_address`、`session_id`、`transition_id` 暂时没有特殊要求时可以留空。

### 3. 安全提醒

`loginConf.ini` 包含后台登录账号、密码、apikey、token 等敏感信息。

请务必妥善保管：

- 不要把 `loginConf.ini` 发给任何人
- 不要把 `loginConf.ini` 提交到 Git 仓库
- 不要截图发送完整配置内容
- 如果怀疑配置泄露，请及时更换后台密码和 apikey

项目已经把 `loginConf.ini` 加入 `.gitignore`，但仍然需要手动注意，不要复制到公开位置。

## 二、配置下载参数

普通接口地址和下载目录在 `config.ini` 的 `[fetch_orders]` 中维护，通常不用修改：

```ini
[fetch_orders]
login_url = https://pgw.jazzcash.com.pk/pgwportal/1.0.0/api/AXAServiceUMG/v1/Auth/login
report_url = https://pgw.jazzcash.com.pk/pgwportal/1.0.0/api/AXAServiceRPT/v1/Reports/getScheduleReady
scheduled_reports_url = https://pgw.jazzcash.com.pk/pgwportal/1.0.0/api/AXAServiceRPT/v1/Reports/getReportsScheduled
download_url = https://pgw.jazzcash.com.pk/pgwportal/1.0.0/api/AXAServiceRPT/v1/Reports/downloadZip
download_dir = workspace/diffOrders
```

默认下载目录是：

```text
workspace/diffOrders
```

下载完成后的文件名类似：

```text
TranDetailReport_87382398_20260629141406.1587508.xlsx
```

## 三、运行方式

### 方式一：开发环境运行

在项目根目录执行：

```bash
PYTHONPATH=src python3 -m autoexcel.fetch_orders
```

### 方式二：dist 打包目录运行

进入打包目录：

```bash
cd dist/AutoExcelKit
./fetch-orders
```

如果是第一次在 `dist/AutoExcelKit` 里运行，请确认该目录下已经存在真实配置：

```text
dist/AutoExcelKit/loginConf.ini
```

注意：项目根目录的 `loginConf.ini` 不会自动复制到 `dist/AutoExcelKit`，需要你手动准备一份真实配置。

## 四、运行时交互说明

程序启动后会先登录后台：

```text
正在登录后台服务，准备拉取订单 Excel...
登录请求完成，HTTP 状态码：200
```

登录成功后，会询问：

```text
是否需要生成最新的 Excel？输入 y 生成，直接回车或输入 n 使用最近已生成文件：
```

### 选择 y

输入 `y` 后，程序会继续询问拉取订单日期：

```text
请输入拉取订单日期，直接回车默认前一天 YYYY-MM-DD：
```

然后调用 `getScheduleReady` 生成指定日期的新报表任务。

日期支持：

```text
2026-06-28
0628
06-28
06/28
```

### 直接回车或输入 n

直接回车或输入 `n` 时，程序不会生成新的 Excel。

它会直接调用 `getReportsScheduled` 获取最近已经生成完成的 Excel 列表，然后选择最新的可下载文件，调用 `downloadZip` 下载。

这是日常测试或已经有报表时推荐的方式。

## 五、下载完成后的结果

成功时会看到类似输出：

```text
已调度报表查询完成，HTTP 状态码：200，记录数：97
准备下载最近生成的 Excel：TranDetailReport_87382398_20260629141406.1587508.xlsx
Excel 下载完成：/path/to/AutoExcelKit/workspace/diffOrders/TranDetailReport_87382398_20260629141406.1587508.xlsx
```

下载完成后，可以继续运行订单对账工具：

```bash
./diff-orders
```

或开发环境运行：

```bash
PYTHONPATH=src python3 -m autoexcel.diff_orders
```

## 六、常见问题

### 1. 登录成功，但查询列表或生成报表返回 401

常见原因是 `loginConf.ini` 里的默认 `api_key` 被误删、改坏，或和当前账号/token 不匹配。

请重点检查：

- `api_key` 是否还保留模板里的默认值
- `username` 是否正确
- `password` 是否正确
- `institution_id` 是否正确
- 当前运行目录下的 `loginConf.ini` 是否是你真正填写的那一份

打包目录运行时，程序读取的是：

```text
dist/AutoExcelKit/loginConf.ini
```

不是项目根目录的 `loginConf.ini`。

### 2. 不想生成新的 Excel，只想下载最近的文件

在提示处直接回车，或输入 `n`：

```text
是否需要生成最新的 Excel？输入 y 生成，直接回车或输入 n 使用最近已生成文件：
```

### 3. 下载后的 Excel 没有出现在对账目录

确认 `config.ini` 中：

```ini
download_dir = workspace/diffOrders
```

并确认程序最后打印的 `Excel 下载完成` 路径。

### 4. 更换账号后仍然异常

删除 `loginConf.ini` 中 `[upstream_server_session]` 下旧的 `token` 等运行时字段，保留 `[upstream_server_login]` 的默认 `api_key` 和三项登录信息，然后重新运行。

程序会重新登录并写入新的 session 信息。
