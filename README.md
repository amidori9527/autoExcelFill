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
that folder and use today's date by default.

Default behavior is controlled by `config.ini` next to the executable. Users can
edit it with a text editor to set `target_date`, `limit_sheets`, `workbook`, and
other fill options. Dates support full or short forms, such as `2026-06-10`,
`0610`, `06-10`, and `05/12`; short forms use the current year.

Before running the fill operation, close the workbook in Excel/WPS. If the
program fails, it writes details to `autoexcel-fill-error.log` next to the
executable.
