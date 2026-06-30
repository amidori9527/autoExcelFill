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

## Diff Orders

Put the upstream and backend workbooks into `workspace/diffOrders`, then run:

```bash
PYTHONPATH=src python3 -m autoexcel.diff_orders
```

The tool asks for a target date, defaults to today, then automatically
selects files by name:

- upstream: `TranDetailReport_<id>_<YYYYMMDD...>.xlsx`
- backend: `收款订单_<YYYYMMDD...>.xlsx`

It compares upstream column L order IDs and H amounts with backend column D
order IDs, G amounts, and I fees, then writes a statistics HTML result into the
`result` folder.

When more than one upstream/backend group matches the date, the tool compares
all groups in batch, writes one summary HTML result, and opens it automatically.
Batch mode does not copy order IDs to the clipboard automatically.

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
