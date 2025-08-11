
import re
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

st.set_page_config(page_title="Weekly Estimated vs Actual Hours", layout="wide")
st.title("📊 Ohio Lottery Weekly Estimated vs Actual Hours Chart")

# ---------------------- Helpers ----------------------
DATE_COL_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")

def normalize_colname(s):
    return str(s).strip().lower()

def detect_project_col(df: pd.DataFrame):
    # Try direct match on columns
    for c in df.columns:
        if normalize_colname(c) == "project full name":
            return c
    # If not found, try to use first row as header if it contains the string
    head_row = df.iloc[0].astype(str).str.strip().str.lower().tolist()
    if "project full name" in head_row:
        df.columns = df.iloc[0]
        df.drop(df.index[0], inplace=True)
        df.reset_index(drop=True, inplace=True)
        for c in df.columns:
            if normalize_colname(c) == "project full name":
                return c
    # As a last resort, assume first column is the project name
    return df.columns[0]

def detect_week_cols(df: pd.DataFrame):
    week_cols = []
    for c in df.columns:
        s = str(c).strip()
        if DATE_COL_PATTERN.match(s):
            week_cols.append(c)
    # Also support datetime-typed columns
    for c in df.columns:
        if isinstance(c, pd.Timestamp):
            if c not in week_cols:
                week_cols.append(c)
    # Ensure left-to-right chronological order
    week_cols_sorted = sorted(week_cols, key=lambda x: pd.to_datetime(str(x)))
    return week_cols_sorted

def get_total_row_estimates(baseline_df: pd.DataFrame, project_col: str, week_cols: list):
    # Find a row whose project name equals 'Total' (case-insensitive, trimmed)
    mask_total = baseline_df[project_col].astype(str).str.strip().str.lower() == "total"
    if mask_total.any():
        total_row = baseline_df.loc[mask_total, week_cols]
        # ensure numeric
        total_row = total_row.apply(pd.to_numeric, errors="coerce")
        sr = total_row.iloc[0]
        return pd.DataFrame({
            "Week": pd.to_datetime([str(c) for c in sr.index]),
            "Estimated_Hours": sr.values
        })
    # Fallback: sum, but inform user via message
    st.warning("Couldn't find a 'Total' row in the baseline; using sum across projects as a fallback.")
    summed = baseline_df[week_cols].apply(pd.to_numeric, errors="coerce").sum()
    return pd.DataFrame({
        "Week": pd.to_datetime([str(c) for c in summed.index]),
        "Estimated_Hours": summed.values
    })

def extract_actual_total_from_file(actuals_df: pd.DataFrame):
    # Robustly try to read the total from the file structure they've been using.
    df = actuals_df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # If there's a clear "Total" row in the first column
    first_col = df.columns[0]
    # Standard two-column format: ["Project Full Name", "Actual Hours Worked"]
    # Try explicit 'Total'
    mask_total = df[first_col].astype(str).str.strip().str.lower() == "total"
    if mask_total.any():
        val = pd.to_numeric(df.loc[mask_total, df.columns[1]].iloc[-1], errors="coerce")
        if pd.notna(val):
            return float(val)
    # If last row has NaN project name with a numeric total in second column
    if df[first_col].isna().any():
        last_nan_row = df[df[first_col].isna()].tail(1)
        if not last_nan_row.empty:
            val = pd.to_numeric(last_nan_row.iloc[0, 1], errors="coerce")
            if pd.notna(val):
                return float(val)
    # Fallback: sum all numeric values in second column
    st.warning("No explicit Total row found in actuals; summing the 'Actual Hours Worked' column as a fallback.")
    col2_numeric = pd.to_numeric(df.iloc[:, 1], errors="coerce")
    return float(col2_numeric.fillna(0).sum())

def one_decimal(df: pd.DataFrame, cols):
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce").round(1)
    return out

# ----------------- Load previous state -----------------
cached_baseline_file = Path("cached_baseline.xlsx")
tracker_file = Path("actuals_tracker.csv")

# Show previous chart & data on startup (if available)
if cached_baseline_file.exists() and tracker_file.exists():
    try:
        prev_base = pd.read_excel(cached_baseline_file, engine="openpyxl")
        project_col_prev = detect_project_col(prev_base)
        week_cols_prev = detect_week_cols(prev_base)
        estimates_prev = get_total_row_estimates(prev_base, project_col_prev, week_cols_prev)

        prev_actuals = pd.read_csv(tracker_file, parse_dates=["Week"])
        combined_prev = pd.merge(estimates_prev, prev_actuals, on="Week", how="left").fillna(0)

        st.subheader("📈 Last Saved View")
        st.caption("Loaded from cached baseline and previously tracked actuals.")
        st.dataframe(
            one_decimal(
                combined_prev.assign(Week=combined_prev["Week"].dt.strftime("%Y-%m-%d")),
                ["Estimated_Hours", "Actual_Hours"]
            )
        )

        fig_prev = go.Figure()
        fig_prev.add_trace(go.Bar(x=combined_prev["Week"], y=combined_prev["Estimated_Hours"], name="Estimated Hours"))
        fig_prev.add_trace(go.Bar(x=combined_prev["Week"], y=combined_prev["Actual_Hours"], name="Actual Hours"))
        fig_prev.update_layout(barmode="overlay", title="Estimated vs Actual Hours per Week (Previous)",
                               xaxis_title="Week", yaxis_title="Hours", height=460)
        st.plotly_chart(fig_prev, use_container_width=True)

        last_html = Path("Estimated_vs_Actual_Hours_Chart.html")
        if last_html.exists():
            with open(last_html, "rb") as f:
                st.download_button("📥 Download Last Chart (HTML)", f, file_name=last_html.name)
    except Exception as e:
        st.info(f"Previous view not available yet ({e}). Upload files below to create the first one.")

st.markdown("---")

# ----------------- Uploaders -----------------
baseline_file = st.file_uploader("Upload Baseline Excel File (optional)", type="xlsx")
actuals_file = st.file_uploader("Upload Actuals Excel File for This Week", type="xlsx")
st.caption("📎 Reminder: Actuals filename should follow this format: **Actuals_YYYY-MM-DD.xlsx**")

# ----------------- Load baseline (optional) -----------------
baseline_df = None
if baseline_file is not None:
    try:
        baseline_df = pd.read_excel(baseline_file, engine="openpyxl")
        baseline_df.to_excel(cached_baseline_file, index=False)
    except Exception as e:
        st.error(f"Could not read baseline file: {e}")
        st.stop()
elif cached_baseline_file.exists():
    baseline_df = pd.read_excel(cached_baseline_file, engine="openpyxl")
else:
    st.warning("Please upload a baseline file for the first use.")
    st.stop()

# ----------------- Process actuals if provided -----------------
if actuals_file is not None:
    st.info("Processing your files...")

    try:
        actuals_df = pd.read_excel(actuals_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Could not read actuals file: {e}")
        st.stop()

    # Detect project column and week columns in baseline
    project_col = detect_project_col(baseline_df)
    week_cols = detect_week_cols(baseline_df)

    if not week_cols:
        st.error("No weekly date columns (YYYY-MM-DD) were found in the baseline file.")
        st.stop()

    # Build estimate series from 'Total' row (preferred)
    estimates = get_total_row_estimates(baseline_df, project_col, week_cols)

    # Extract this week's actual total value from the uploaded file
    actual_value = extract_actual_total_from_file(actuals_df)

    # Get the Week date from the uploaded file name
    fname = actuals_file.name
    m = re.search(r"(\d{4}-\d{2}-\d{2})", fname)
    if not m:
        st.error(f"Filename must include a date in YYYY-MM-DD format. Got: {fname}")
        st.stop()
    actual_week_date = pd.to_datetime(m.group(1))

    # Load or create tracker and show preview
    if tracker_file.exists():
        actuals_tracker = pd.read_csv(tracker_file, parse_dates=["Week"])
        st.subheader("📆 Previous Weeks' Actuals")
        st.dataframe(one_decimal(actuals_tracker.assign(Week=actuals_tracker["Week"].dt.strftime("%Y-%m-%d")), ["Actual_Hours"]))
    else:
        actuals_tracker = pd.DataFrame(columns=["Week", "Actual_Hours"])

    # Append/update this week's actual
    new_row = pd.DataFrame({"Week": [actual_week_date], "Actual_Hours": [actual_value]})
    actuals_tracker = pd.concat([actuals_tracker, new_row], ignore_index=True)
    actuals_tracker.drop_duplicates(subset=["Week"], keep="last", inplace=True)
    actuals_tracker.sort_values("Week", inplace=True)
    actuals_tracker.to_csv(tracker_file, index=False)
    st.success("✅ Actuals tracker updated!")

    # Merge for chart
    combined = pd.merge(estimates, actuals_tracker, on="Week", how="left").fillna(0)

    # Plot
    fig = go.Figure()
    fig.add_trace(go.Bar(x=combined["Week"], y=combined["Estimated_Hours"], name="Estimated Hours", marker_color="lightblue"))
    fig.add_trace(go.Bar(x=combined["Week"], y=combined["Actual_Hours"], name="Actual Hours", marker_color="darkblue"))
    fig.update_layout(barmode="overlay", title="Estimated vs Actual Hours per Week",
                      xaxis_title="Week", yaxis_title="Hours", height=520)
    st.success("✅ Chart generated!")
    st.plotly_chart(fig, use_container_width=True)

    # Save downloadable HTML
    with open("Estimated_vs_Actual_Hours_Chart.html", "w") as f:
        pio.write_html(fig, file=f, auto_open=False)
    with open("Estimated_vs_Actual_Hours_Chart.html", "rb") as f:
        st.download_button("📥 Download Chart as HTML", f, file_name="Estimated_vs_Actual_Hours_Chart.html")

else:
    st.info("Upload this week's Actuals file to generate the latest chart.")
