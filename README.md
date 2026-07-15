# autoexcel

Python scripts for automating operations on `data.xlsx`.

## Run

```bash
PYTHONPATH=src python3 -m autoexcel.main
```

List sheets:

```bash
PYTHONPATH=src python3 -m autoexcel.main --list-sheets
```

Preview another sheet:

```bash
PYTHONPATH=src python3 -m autoexcel.main --sheet Sheet1
```

If running from a clean environment:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
PYTHONPATH=src .venv/bin/python -m autoexcel.main
```

## Packaged app

Put the target `.xlsx` file into the `workspace` folder next to the packaged
executable, then run the executable directly. The tool will list workbooks in
that folder and ask whether to use the default date or enter a date manually.
The fill tool's default date remains based on its existing logic.

Default behavior is controlled by `config.ini` next to the executable. Users can
edit it with a text editor to set `target_date`, `limit_sheets`, `workbook`, and
other fill options. Dates support full or short forms, such as `2026-06-10`,
`0610`, `06-10`, and `05/12`; short forms use the current year.

Before running the fill operation, close the workbook in Excel/WPS. If the
program fails, it writes details to `autoexcel-fill-error.log` next to the
executable. Successful runs write detailed processing logs to the `logs` folder
next to the executable.

## Desktop demo

Run the PySide6 desktop interface from source:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
PYTHONPATH=src .venv/bin/python -m autoexcel.gui
```

Build the desktop application on the current operating system:

```bash
.venv/bin/pyinstaller -y autoexcel-gui.spec
```

The macOS build is written to `dist/AutoExcel.app`. Windows produces the
`dist/AutoExcel` application directory. Build separately on each operating
system; PyInstaller does not produce a Windows executable from macOS.

### Build the Windows toolkit

Use 64-bit Python 3.10 or newer on Windows. From PowerShell in the repository
root, create an isolated environment and run the packaging script:

```powershell
py -3.10 -m venv .venv
Set-ExecutionPolicy -Scope Process Bypass
.\package-windows.ps1 -Python ".\.venv\Scripts\python.exe"
```

The script installs dependencies, runs the test suite, builds all five command
line tools, verifies that every expected `.exe` exists, and creates
`AutoExcelKit-Windows.zip`. The archive includes launchers for fill, order
comparison, order download, adding cards, and adding B2B data. Do not add real
workbooks or `loginConf.ini` to the repository before building.

The desktop interface includes a visual editor for the global `config.ini`.
The order comparison page can also edit the selected group's `conf.ini`
platform setting. Login credentials and session values remain isolated in
`loginConf.ini` and are never displayed by the settings page.

### Feature licenses

Without a valid license, the desktop interface exposes only Excel fill and
settings. Order comparison and order download are enabled by an Ed25519-signed
license key entered in the settings page. The license is stored in `license.key`
beside the distributed application and is not written to `config.ini`.

Generate a permanent license for both protected features:

```bash
PYTHONPATH=src .venv/bin/python -m autoexcel.license_tool \
  --output /tmp/autoexcel-customer.key \
  --customer customer-name
```

Add `--expires 2027-12-31` for a license that expires at the end of that UTC
date. Use `--features order_diff` or `--features fetch_orders` for a partial
license. The signing key is stored locally at
`.license/autoexcel-ed25519-private.pem`, is ignored by Git, and must be backed
up securely. It is never included in PyInstaller output; only the public
verification key is embedded in the application.

## Diff Orders

Put the upstream and backend workbooks into `workspace/diffOrders`, then run:

```bash
PYTHONPATH=src python3 -m autoexcel.diff_orders
```

The tool asks for a target date, defaults to today, then automatically
selects files by name:

- upstream: `TranDetailReport_<id>_<YYYYMMDD...>.xlsx` or `PGW_TXNDETAIL_<id>_<YYYYMMDD>_....xlsx`
- backend: `收款订单_<YYYYMMDD...>.xlsx`

It compares upstream column L order IDs and H amounts with backend column D
order IDs, G amounts, and I fees, then writes a statistics HTML result into the
`result` folder.

When more than one upstream/backend group matches the date, the tool compares
all groups in batch, writes one summary HTML result, and opens it automatically.
Batch mode does not copy order IDs to the clipboard automatically.

When a group contains `代收重复支付订单_<date>.xlsx`, the tool first compares
upstream vs TP, then compares the resulting difference order IDs with duplicate
payment column C to split duplicate orders from remaining differences.
Put `conf.ini` in a comparison group folder to choose the algorithm:
`[diff_orders] platform = finerbit` or `[diff_orders] platform = easypaisa`.
If `conf.ini` is missing, unreadable, or does not define `platform`,
duplicate-payment groups default to `finerbit`.
For this finerBit mode, the upstream fee uses upstream column G amount and
upstream column D ChannelName: EasyPaisa 4%, JazzCash 2.3%.
The finerBit channel cost uses the same upstream amount and channel, rounded to
2 decimal places per order before summing.
For easypaisa mode, the upstream fee is upstream columns R + S, and the channel
cost is TP order amount * 0.02.

Set `[diff_orders] auto_open_html = false` in `config.ini` to stop automatically
opening the generated HTML. Result HTML filenames include the comparison folder
name, for example `order_diff_jz663_YYYYMMDD_HHMMSS.html`.

## Fetch Orders

Fill non-sensitive `[fetch_orders]` settings in `config.ini`. Copy
`loginConf.example.ini` to `loginConf.ini`, fill `[upstream_server_login]`, then run:

```bash
PYTHONPATH=src python3 -m autoexcel.fetch_orders
```

After login, the command asks whether to generate a fresh Excel report. Choose
`y` to call `getScheduleReady`; press Enter or choose `n` to skip generation,
enter a download date that defaults to yesterday, then download the first
available scheduled report for that date through `downloadZip` into
`workspace/diffOrders`.

It does not print sensitive values to the console; the upstream login token,
runtime session values, and last downloaded file metadata are written back to
`[upstream_server_session]` in `loginConf.ini`.
`loginConf.ini` is ignored by git.
