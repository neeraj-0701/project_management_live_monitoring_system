# R&D Lab Management Dashboard

An interactive Streamlit dashboard built from `Project_Management__1_.xlsx`.

## Setup (one-time)

```bash
pip install -r requirements.txt
```

## Run

Place `Project_Management__1_.xlsx` in the same folder as `app.py` (already done), then:

```bash
streamlit run app.py
```

This opens the dashboard in your browser at `http://localhost:8501`.

## Updating data

Two ways to refresh the dashboard after you fill in more rows in Excel:

1. **Simplest:** overwrite `Project_Management__1_.xlsx` in this folder with your
   updated file, then refresh the browser tab (Streamlit reloads automatically).
2. **No overwrite needed:** use the "Upload workbook" control in the sidebar to
   point the dashboard at any copy of the file, without touching the code.

## What it covers

| Tab | Source sheet | Shows |
|---|---|---|
| Materials | Materials & Inventory | Stock levels, critical/low-stock items, reorder list |
| HR | Human Resources | Task status breakdown, tasks per employee |
| Casting & Testing | Casting and Testing Activity | Mix status, target values per mix |
| Flow Test | Flow Table Test | Flow (mm) trend over time |
| Compression Test | Compression Test | Compressive strength vs target |
| Flexural Test | Flexural Test | Flexural strength trend over time |
| Budget | Budget | Cost by head, allocation split, total budget KPI |
| SCM | SCM | Purchase status, spend by item |

## Why it looks sparse right now

Most sheets in the current workbook are templates with only a handful of rows
filled in (Materials, Budget and SCM have real data; the test-log sheets only
have dates typed in so far). The dashboard is built to auto-detect each
sheet's header row and skip empty sections gracefully — as you log more test
results and HR tasks, those charts will populate on their own without any
code changes.

## Customizing

- **Add a chart:** each tab is a self-contained block in `app.py` using
  `plotly.express` — copy an existing `px.bar(...)` / `px.line(...)` /
  `px.pie(...)` call and point it at different columns.
- **Add filters:** `st.sidebar.multiselect` / `st.sidebar.date_input` can be
  used to filter any of the loaded DataFrames before plotting.
- **Change colors/theme:** create a `.streamlit/config.toml` with a
  `[theme]` section, or pass `color_discrete_sequence=[...]` to any Plotly call.
