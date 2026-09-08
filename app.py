"""
R&D Lab Management Dashboard
=============================
Reads the "Project_Management" Excel workbook and renders an interactive
Streamlit dashboard covering Materials & Inventory, HR tasks, Casting &
Testing activity, Flow/Compression/Flexural test logs, Budget and SCM
purchasing.

Run with:
    pip install streamlit pandas openpyxl plotly
    streamlit run app.py

By default it looks for "Project_Management__1_.xlsx" in the same folder.
You can also upload a different (or updated) copy of the workbook from the
sidebar without touching the code.
"""

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="R&D Lab Dashboard", layout="wide", page_icon="🧪")

DEFAULT_FILE = "Project_Management__1_.xlsx"

# ----------------------------------------------------------------------
# Generic helpers
# ----------------------------------------------------------------------

def find_header_row(raw: pd.DataFrame, marker: str, search_rows: int = 20) -> int:
    """Scan the first `search_rows` rows for one containing `marker` in any
    cell, and return that row's index. Falls back to 0 if not found."""
    marker = marker.strip().lower()
    for i in range(min(search_rows, len(raw))):
        row_vals = raw.iloc[i].astype(str).str.strip().str.lower()
        if (row_vals == marker).any():
            return i
    return 0


def load_sheet(xls: pd.ExcelFile, sheet_name: str, marker: str) -> pd.DataFrame:
    """Load a sheet, auto-detecting its real header row via `marker`
    (e.g. 'S.No' or 'Cost Head'), then drop fully-empty rows/columns."""
    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    header_row = find_header_row(raw, marker)
    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")
    # Drop the stray "Unnamed" columns Excel leaves for merged/blank cells
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    return df.reset_index(drop=True)


def has_data(df: pd.DataFrame, min_rows: int = 1) -> bool:
    return df is not None and not df.empty and len(df) >= min_rows


def kpi(col, label, value, help_text=None):
    col.metric(label, value, help=help_text)


def empty_notice(section: str):
    st.info(f"No {section} data entered yet — this panel will populate automatically once rows are filled in the sheet.")


# ----------------------------------------------------------------------
# Sidebar: file source
# ----------------------------------------------------------------------

st.sidebar.title("🧪 R&D Lab Dashboard")
uploaded = st.sidebar.file_uploader("Upload workbook (.xlsx)", type=["xlsx"])
source = uploaded if uploaded is not None else DEFAULT_FILE

try:
    xls = pd.ExcelFile(source)
except FileNotFoundError:
    st.error(
        f"Could not find '{DEFAULT_FILE}'. Place it next to app.py, "
        "or upload a workbook using the sidebar."
    )
    st.stop()

st.sidebar.caption("Data source: " + (uploaded.name if uploaded else DEFAULT_FILE))

# ----------------------------------------------------------------------
# Load each sheet (resilient to sparse/template data)
# ----------------------------------------------------------------------

materials = load_sheet(xls, "Materials & Inventory", "S.No") if "Materials & Inventory" in xls.sheet_names else pd.DataFrame()
hr = load_sheet(xls, "Human Resources ", "S.No") if "Human Resources " in xls.sheet_names else pd.DataFrame()
casting = load_sheet(xls, "Casting and Testing Activity", "S.No") if "Casting and Testing Activity" in xls.sheet_names else pd.DataFrame()
flow = load_sheet(xls, "Flow Table Test ", "S.No") if "Flow Table Test " in xls.sheet_names else pd.DataFrame()
compression = load_sheet(xls, "Compression Test ", "S.No") if "Compression Test " in xls.sheet_names else pd.DataFrame()
flexural = load_sheet(xls, "Flexural  Test ", "S.No") if "Flexural  Test " in xls.sheet_names else pd.DataFrame()
budget = load_sheet(xls, "Budget", "Cost Head") if "Budget" in xls.sheet_names else pd.DataFrame()
scm = load_sheet(xls, "SCM", "S.No") if "SCM" in xls.sheet_names else pd.DataFrame()

# Drop rows that only have an S.No / Date but nothing else meaningful
def drop_blank_records(df, key_cols):
    present = [c for c in key_cols if c in df.columns]
    if not present:
        return df
    check_cols = [c for c in df.columns if c not in ("S.No", "Date")]
    if not check_cols:
        return df
    return df[df[check_cols].notna().any(axis=1)].reset_index(drop=True)


materials = drop_blank_records(materials, ["S.No", "Date"])
casting = drop_blank_records(casting, ["S.No", "Date"])
flow = drop_blank_records(flow, ["S.No", "Date"])
compression = drop_blank_records(compression, ["S.No", "Date"])
flexural = drop_blank_records(flexural, ["S.No", "Date"])

# ----------------------------------------------------------------------
# Top-level KPIs
# ----------------------------------------------------------------------

st.title("R&D Lab Management Dashboard")
st.caption("Materials • Testing • Casting • Resources • Budget")

c1, c2, c3, c4, c5 = st.columns(5)
unique_materials = materials["Material"].nunique() if has_data(materials) and "Material" in materials.columns else 0
kpi(c1, "Total Materials", unique_materials)
critical_count = int((materials.get("Stock Status", pd.Series(dtype=str)) == "Critical").sum()) if has_data(materials) else 0
kpi(c2, "Critical Materials", critical_count)
kpi(c3, "Casting Activities", len(casting) if has_data(casting) else 0)
total_tests = sum(has_data(d) for d in (flow, compression, flexural) if True) and (
    (len(flow) if has_data(flow) else 0) + (len(compression) if has_data(compression) else 0) + (len(flexural) if has_data(flexural) else 0)
)
kpi(c4, "Total Tests Logged", total_tests)
if has_data(budget) and "Basic Amount" in budget.columns:
    total_row = budget[budget.iloc[:, 0].astype(str).str.contains("total", case=False, na=False)]
    total_budget = total_row["Basic Amount"].sum() if not total_row.empty else budget["Basic Amount"].sum()
else:
    total_budget = 0
kpi(c5, "Total Budget (₹)", f"{total_budget:,.0f}" if total_budget else "—")

st.divider()

tabs = st.tabs(["📦 Materials", "👥 HR", "🏗️ Casting & Testing", "💧 Flow Test", "🧱 Compression Test", "📐 Flexural Test", "💰 Budget", "🛒 SCM"])

# ----------------------------------------------------------------------
# Materials & Inventory
# ----------------------------------------------------------------------
with tabs[0]:
    st.subheader("Materials & Inventory")
    if has_data(materials):
        colA, colB = st.columns([2, 1])
        with colA:
            if {"Material", "Closing Stock", "Stock Status"}.issubset(materials.columns):
                fig = px.bar(
                    materials, x="Material", y="Closing Stock", color="Stock Status",
                    color_discrete_map={"Critical": "#e74c3c", "Adequate": "#2ecc71"},
                    title="Closing Stock by Material",
                )
                st.plotly_chart(fig, use_container_width=True)
        with colB:
            if "Stock Status" in materials.columns:
                counts = materials["Stock Status"].value_counts().reset_index()
                counts.columns = ["Stock Status", "Count"]
                fig2 = px.pie(counts, names="Stock Status", values="Count", title="Stock Status Mix",
                               color="Stock Status", color_discrete_map={"Critical": "#e74c3c", "Adequate": "#2ecc71"})
                st.plotly_chart(fig2, use_container_width=True)
        if "Purchase Required" in materials.columns:
            reorder = materials[materials["Purchase Required"].astype(str).str.lower() == "yes"]
            if not reorder.empty:
                st.markdown("**Items needing reorder**")
                st.dataframe(reorder, use_container_width=True, hide_index=True)
        with st.expander("Full materials table"):
            st.dataframe(materials, use_container_width=True, hide_index=True)
    else:
        empty_notice("materials")

# ----------------------------------------------------------------------
# HR
# ----------------------------------------------------------------------
with tabs[1]:
    st.subheader("Human Resources")
    if has_data(hr):
        name_col = "Employee Name" if "Employee Name" in hr.columns else hr.columns[0]
        st.markdown(f"**Team members logged:** {hr[name_col].nunique()}")
        if "Task Status" in hr.columns and hr["Task Status"].notna().any():
            fig = px.pie(hr, names="Task Status", title="Task Status Distribution")
            st.plotly_chart(fig, use_container_width=True)
        if "Task" in hr.columns and hr["Task"].notna().any():
            task_counts = hr.groupby(name_col)["Task"].count().reset_index(name="Tasks")
            fig2 = px.bar(task_counts, x=name_col, y="Tasks", title="Tasks per Employee")
            st.plotly_chart(fig2, use_container_width=True)
        with st.expander("Full HR log"):
            st.dataframe(hr, use_container_width=True, hide_index=True)
    else:
        empty_notice("HR")

# ----------------------------------------------------------------------
# Casting & Testing Activity
# ----------------------------------------------------------------------
with tabs[2]:
    st.subheader("Casting & Testing Activity")
    if has_data(casting):
        if "Status" in casting.columns and casting["Status"].notna().any():
            fig = px.pie(casting, names="Status", title="Mix Status")
            st.plotly_chart(fig, use_container_width=True)
        target_cols = [c for c in ["Target Flow (mm)", "Target Compressive Strength", "Target Flexural Strength"] if c in casting.columns]
        if target_cols and "Mix ID" in casting.columns:
            melted = casting.melt(id_vars=["Mix ID"], value_vars=target_cols, var_name="Target", value_name="Value")
            melted = melted.dropna(subset=["Value"])
            if not melted.empty:
                fig2 = px.bar(melted, x="Mix ID", y="Value", color="Target", barmode="group", title="Targets by Mix")
                st.plotly_chart(fig2, use_container_width=True)
        with st.expander("Full casting log"):
            st.dataframe(casting, use_container_width=True, hide_index=True)
    else:
        empty_notice("casting & testing")

# ----------------------------------------------------------------------
# Flow Table Test
# ----------------------------------------------------------------------
with tabs[3]:
    st.subheader("Flow Table Test")
    if has_data(flow):
        if {"Date", "Flow Table  (mm)"}.issubset(flow.columns):
            fig = px.line(flow.sort_values("Date"), x="Date", y="Flow Table  (mm)",
                           color="Mix ID" if "Mix ID" in flow.columns else None,
                           markers=True, title="Flow Table Result Over Time")
            st.plotly_chart(fig, use_container_width=True)
        with st.expander("Full flow test log"):
            st.dataframe(flow, use_container_width=True, hide_index=True)
    else:
        empty_notice("flow table test")

# ----------------------------------------------------------------------
# Compression Test
# ----------------------------------------------------------------------
with tabs[4]:
    st.subheader("Compression Test")
    if has_data(compression):
        if {"Compressive Strength Mpa", "Target Strength Mpa"}.issubset(compression.columns):
            plot_df = compression.dropna(subset=["Compressive Strength Mpa"])
            if not plot_df.empty:
                fig = px.bar(plot_df, x="Specimen ID" if "Specimen ID" in plot_df.columns else plot_df.index,
                             y=["Compressive Strength Mpa", "Target Strength Mpa"],
                             barmode="group", title="Compressive Strength vs Target")
                st.plotly_chart(fig, use_container_width=True)
        with st.expander("Full compression test log"):
            st.dataframe(compression, use_container_width=True, hide_index=True)
    else:
        empty_notice("compression test")

# ----------------------------------------------------------------------
# Flexural Test
# ----------------------------------------------------------------------
with tabs[5]:
    st.subheader("Flexural Test")
    if has_data(flexural):
        if {"Date", "Flexuarl Strength (Mpa)"}.issubset(flexural.columns):
            plot_df = flexural.dropna(subset=["Flexuarl Strength (Mpa)"])
            if not plot_df.empty:
                fig = px.line(plot_df.sort_values("Date"), x="Date", y="Flexuarl Strength (Mpa)",
                               color="Mix ID" if "Mix ID" in plot_df.columns else None,
                               markers=True, title="Flexural Strength Over Time")
                st.plotly_chart(fig, use_container_width=True)
        with st.expander("Full flexural test log"):
            st.dataframe(flexural, use_container_width=True, hide_index=True)
    else:
        empty_notice("flexural test")

# ----------------------------------------------------------------------
# Budget
# ----------------------------------------------------------------------
with tabs[6]:
    st.subheader("Budget")
    if has_data(budget):
        cost_col = budget.columns[0]
        line_items = budget[~budget[cost_col].astype(str).str.contains("total|sub", case=False, na=False)]
        line_items = line_items.dropna(subset=["Basic Amount"]) if "Basic Amount" in line_items.columns else line_items
        colA, colB = st.columns(2)
        with colA:
            if not line_items.empty and "Basic Amount" in line_items.columns:
                fig = px.bar(line_items, x=cost_col, y="Basic Amount", title="Cost by Head (Basic Amount, ₹)")
                st.plotly_chart(fig, use_container_width=True)
        with colB:
            if not line_items.empty and "Basic Amount" in line_items.columns:
                fig2 = px.pie(line_items, names=cost_col, values="Basic Amount", title="Budget Allocation")
                st.plotly_chart(fig2, use_container_width=True)
        with st.expander("Full budget table"):
            st.dataframe(budget, use_container_width=True, hide_index=True)
    else:
        empty_notice("budget")

# ----------------------------------------------------------------------
# SCM
# ----------------------------------------------------------------------
with tabs[7]:
    st.subheader("SCM — Equipment & Instruments Procurement")
    if has_data(scm):
        if "Purchase Status" in scm.columns:
            fig = px.pie(scm, names="Purchase Status", title="Purchase Status")
            st.plotly_chart(fig, use_container_width=True)
        if {"Item Description", "Amount (₹)"}.issubset(scm.columns):
            amt = scm.copy()
            amt["Amount (₹)"] = pd.to_numeric(amt["Amount (₹)"], errors="coerce")
            amt = amt.dropna(subset=["Amount (₹)"])
            if not amt.empty:
                fig2 = px.bar(amt, x="Item Description", y="Amount (₹)", title="Spend by Item")
                fig2.update_xaxes(tickangle=45)
                st.plotly_chart(fig2, use_container_width=True)
        with st.expander("Full SCM table"):
            st.dataframe(scm, use_container_width=True, hide_index=True)
    else:
        empty_notice("SCM")
