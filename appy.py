import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Submission Dashboard", layout="wide")

# -------------------------
# SESSION STATE
# -------------------------
if "uploaded_data" not in st.session_state:
    st.session_state.uploaded_data = None

if "file_type" not in st.session_state:
    st.session_state.file_type = None

if "anchor_dates" not in st.session_state:
    st.session_state.anchor_dates = pd.DataFrame(
        columns=["Anchor Date", "Date", "Status"]
    )

# -------------------------
# HELPER FUNCTIONS
# -------------------------
def read_submission_excel(uploaded_file, sheet_name):
    raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

    header_row = None

    for i, row in raw.iterrows():
        row_text = row.astype(str).str.strip().str.lower().tolist()

        if "task name" in row_text and (
            "planned start" in row_text or "planned finish" in row_text
        ):
            header_row = i
            break

    if header_row is not None:
        headers = raw.iloc[header_row].fillna("").astype(str).str.strip()

        clean_headers = []
        for idx, header in enumerate(headers):
            if header == "" or header.lower() == "nan":
                clean_headers.append(f"blank_col_{idx}")
            else:
                clean_headers.append(header)

        df = raw.iloc[header_row + 1:].copy()
        df.columns = clean_headers
    else:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )

    df = df.dropna(how="all")

    rename_map = {}

    for col in df.columns:
        c = col.strip().lower()

        if c == "actual start":
            rename_map[col] = "Actual Start"
        elif c == "actual finish":
            rename_map[col] = "Actual Finish"
        elif c == "planned start":
            rename_map[col] = "Planned Start"
        elif c == "planned finish":
            rename_map[col] = "Planned Finish"
        elif c == "task name":
            rename_map[col] = "Task Name"
        elif c == "component id":
            rename_map[col] = "Component ID"
        elif c == "component source":
            rename_map[col] = "Component Source"
        elif c == "filing status":
            rename_map[col] = "Filing Status"
        elif c == "task index":
            rename_map[col] = "Task Index"
        elif c == "wave":
            rename_map[col] = "Wave"
        elif c == "module":
            rename_map[col] = "Module"

    df = df.rename(columns=rename_map)

    # Try to find Wave if missing
    if "Wave" not in df.columns:
        for col in df.columns:
            if df[col].astype(str).str.contains("Rolling Submission|Wave", case=False, na=False).any():
                df["Wave"] = df[col].ffill()
                break

    # Try to find Module if missing
    if "Module" not in df.columns:
        for col in df.columns:
            if df[col].astype(str).str.contains("Module", case=False, na=False).any():
                df["Module"] = df[col].ffill()
                break

    return df


def clean_submission_data(df):
    df.columns = df.columns.astype(str).str.strip()

    required_cols = [
        "Task Name",
        "Planned Start",
        "Actual Start",
        "Planned Finish",
        "Actual Finish"
    ]

    missing = [col for col in required_cols if col not in df.columns]

    if missing:
        st.error(f"Missing required columns: {missing}")
        st.write("Columns found:", list(df.columns))
        st.stop()

    # Keep only actual task rows
    if "Component ID" in df.columns:
        df = df[df["Component ID"].notna() & df["Task Name"].notna()]
    else:
        df = df[df["Task Name"].notna()]

    for col in ["Planned Start", "Actual Start", "Planned Finish", "Actual Finish"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    if "Filing Status" in df.columns:
        df["Filing Status"] = df["Filing Status"].fillna("Incomplete")
    else:
        df["Filing Status"] = df["Actual Finish"].apply(
            lambda x: "Completed" if pd.notna(x) else "Incomplete"
        )

    # Power BI logic
    if "Actually Completed" in df.columns:
        df["Actually Completed"] = (
            df["Actually Completed"]
            .astype(str)
            .str.strip()
            .str.lower()
            .map({"true": True, "false": False})
        )

        df["Actually Completed"] = df["Actually Completed"].fillna(
            df["Actual Finish"].notna()
        )
    else:
        df["Actually Completed"] = df["Actual Finish"].notna()

    if "Planned Completed" in df.columns:
        df["Planned Completed"] = (
            df["Planned Completed"]
            .astype(str)
            .str.strip()
            .str.lower()
            .map({"true": True, "false": False})
        )

        df["Planned Completed"] = df["Planned Completed"].fillna(
            df["Planned Finish"].notna()
        )
    else:
        df["Planned Completed"] = df["Planned Finish"].notna()

    df["StartVarianceDays"] = (
        df["Actual Start"] - df["Planned Start"]
    ).dt.days

    df["FinishVarianceDays"] = (
        df["Actual Finish"] - df["Planned Finish"]
    ).dt.days

    if "Wave" not in df.columns:
        df["Wave"] = "No Wave Found"

    if "Module" not in df.columns:
        df["Module"] = "No Module Found"

    return df


def calculate_metrics(df):
    total_records = len(df)
    completed = df["Actually Completed"].eq(True).sum()
    incomplete = df["Actually Completed"].eq(False).sum()
    planned_completed = df["Planned Completed"].eq(True).sum()

    completion_rate = completed / total_records if total_records > 0 else 0

    finish_variance_sum = 0
    if "FinishVarianceDays" in df.columns:
        finish_variance_sum = df["FinishVarianceDays"].sum(skipna=True)

    return total_records, completed, incomplete, planned_completed, completion_rate, finish_variance_sum


def get_wave_summary(df):
    wave_summary = df.groupby("Wave").agg(
        Total_Records=("Task Name", "count"),
        Actually_Completed_Count=("Actually Completed", lambda x: x.eq(True).sum()),
        Incomplete_Count=("Actually Completed", lambda x: x.eq(False).sum()),
        Planned_Completed_Count=("Planned Completed", lambda x: x.eq(True).sum())
    ).reset_index()

    wave_summary["Completion_Rate"] = (
        wave_summary["Actually_Completed_Count"] /
        wave_summary["Total_Records"] * 100
    )

    wave_summary["Actual_Percent"] = (
        wave_summary["Actually_Completed_Count"] /
        wave_summary["Planned_Completed_Count"] * 100
    ).fillna(0)

    wave_summary["Actual_Percent"] = wave_summary["Actual_Percent"].clip(upper=100)
    wave_summary["Planned_Remaining_Percent"] = 100 - wave_summary["Actual_Percent"]

    return wave_summary


def get_module_summary(df):
    module_summary = df.groupby("Module").agg(
        Total_Records=("Task Name", "count"),
        Actually_Completed_Count=("Actually Completed", lambda x: x.eq(True).sum()),
        Incomplete_Count=("Actually Completed", lambda x: x.eq(False).sum()),
        Planned_Completed_Count=("Planned Completed", lambda x: x.eq(True).sum())
    ).reset_index()

    module_summary["Completion_Rate"] = (
        module_summary["Actually_Completed_Count"] /
        module_summary["Total_Records"] * 100
    )

    return module_summary


def require_upload():
    if st.session_state.uploaded_data is None:
        st.warning("Please upload an Excel file first from the Upload Excel page.")
        return False
    return True


# -------------------------
# SIDEBAR
# -------------------------
st.sidebar.title("Navigation")

page = st.sidebar.radio(
    "Go to",
    [
        "Upload Excel",
        "Rolling Wave View",
        "Rolling Progress View",
        "Drill-Through Component Detail",
        "Non-Rolling Module View",
        "Anchor Dates"
    ]
)

# -------------------------
# UPLOAD EXCEL PAGE
# -------------------------
if page == "Upload Excel":
    st.title("Upload Excel File")
    st.write("Upload the rolling or non-rolling Excel file for processing.")

    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=["xlsx", "xls"]
    )

    if uploaded_file is not None:
        try:
            file_name = uploaded_file.name.lower()

            if "non" in file_name:
                st.session_state.file_type = "Non-Rolling"
            elif "rolling" in file_name:
                st.session_state.file_type = "Rolling"
            else:
                st.session_state.file_type = "Unknown"

            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names

            st.success("File uploaded successfully.")
            st.write("Sheets found:", sheet_names)
            st.write("Detected file type:", st.session_state.file_type)

            selected_sheet = st.selectbox("Select sheet", sheet_names)

            df = read_submission_excel(uploaded_file, selected_sheet)
            df = clean_submission_data(df)

            st.session_state.uploaded_data = df

            st.subheader("Preview")
            st.dataframe(df, use_container_width=True)

            st.subheader("Detected Columns")
            st.write(list(df.columns))

            total, completed, incomplete, planned_completed, completion_rate, finish_variance_sum = calculate_metrics(df)

            st.subheader("Validation Metrics")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Records", total)
            c2.metric("Completed", completed)
            c3.metric("Incomplete", incomplete)
            c4.metric("Completion Rate", f"{completion_rate:.2%}")

        except Exception as e:
            st.error(f"Error reading Excel file: {e}")

# -------------------------
# ROLLING WAVE VIEW
# -------------------------
elif page == "Rolling Wave View":
    st.title("Rolling Submission – Wave View")

    if require_upload():
        df = st.session_state.uploaded_data

        if "Wave" not in df.columns:
            st.error("The uploaded file does not contain a Wave column.")
        else:
            wave_summary = get_wave_summary(df)

            st.subheader("Wave Summary")
            st.dataframe(
                wave_summary[
                    [
                        "Wave",
                        "Actually_Completed_Count",
                        "Incomplete_Count",
                        "Planned_Completed_Count",
                        "Total_Records",
                        "Completion_Rate"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )

            col1, col2 = st.columns(2)

            with col1:
                fig_completion = px.bar(
                    wave_summary,
                    x="Wave",
                    y="Completion_Rate",
                    title="Completion Rate by Wave",
                    text="Completion_Rate"
                )
                fig_completion.update_traces(
                    texttemplate="%{text:.2f}%",
                    textposition="outside"
                )
                fig_completion.update_layout(
                    yaxis_title="Completion Rate (%)",
                    xaxis_title="Wave",
                    yaxis_range=[0, 100]
                )
                st.plotly_chart(fig_completion, use_container_width=True)

            with col2:
                fig_incomplete = px.bar(
                    wave_summary,
                    x="Wave",
                    y="Incomplete_Count",
                    title="Incomplete Documents by Wave",
                    text="Incomplete_Count"
                )
                fig_incomplete.update_layout(
                    yaxis_title="Incomplete Count",
                    xaxis_title="Wave"
                )
                st.plotly_chart(fig_incomplete, use_container_width=True)

            percent_df = wave_summary.melt(
                id_vars="Wave",
                value_vars=[
                    "Actual_Percent",
                    "Planned_Remaining_Percent"
                ],
                var_name="Completion Type",
                value_name="Percent"
            )

            fig_planned_actual = px.bar(
                percent_df,
                y="Wave",
                x="Percent",
                color="Completion Type",
                orientation="h",
                title="Actually Completed Count and Planned Completed Count by Wave",
                barmode="stack"
            )

            fig_planned_actual.update_layout(
                xaxis_title="Actually Completed Count and Planned Completed Count",
                yaxis_title="Wave",
                xaxis_range=[0, 100]
            )

            st.plotly_chart(fig_planned_actual, use_container_width=True)

# -------------------------
# ROLLING PROGRESS VIEW
# -------------------------
elif page == "Rolling Progress View":
    st.title("Rolling Submission – Progress View")

    if require_upload():
        df = st.session_state.uploaded_data

        total, completed, incomplete, planned_completed, completion_rate, finish_variance_sum = calculate_metrics(df)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Completed Documents", completed)
        col2.metric("Total Documents", total)
        col3.metric("Remaining Documents", incomplete)
        col4.metric("Avg Finish Variance (Days)", f"{finish_variance_sum:,.0f}")

        st.divider()

        col_a, col_b = st.columns(2)

        with col_a:
            if "Filing Status" in df.columns:
                status_df = (
                    df["Filing Status"]
                    .fillna("Unknown")
                    .value_counts()
                    .reset_index()
                )
                status_df.columns = ["Filing Status", "Count"]

                fig_donut = px.pie(
                    status_df,
                    names="Filing Status",
                    values="Count",
                    hole=0.45,
                    title="Document Status Breakdown"
                )
                st.plotly_chart(fig_donut, use_container_width=True)
            else:
                st.warning("Filing Status column not found.")

        with col_b:
            fig_gauge = go.Figure(
                go.Indicator(
                    mode="gauge+number",
                    value=completion_rate * 100,
                    number={"suffix": "%", "valueformat": ".2f"},
                    title={"text": "Overall Completion Rate"},
                    gauge={
                        "axis": {"range": [0, 100]},
                        "bar": {"color": "#4285F4"}
                    }
                )
            )
            st.plotly_chart(fig_gauge, use_container_width=True)

        if "Wave" in df.columns:
            wave_summary = get_wave_summary(df)

            percent_df = wave_summary.melt(
                id_vars="Wave",
                value_vars=[
                    "Planned_Remaining_Percent",
                    "Actual_Percent"
                ],
                var_name="Progress Type",
                value_name="Percent"
            )

            fig_progress = px.bar(
                percent_df,
                x="Wave",
                y="Percent",
                color="Progress Type",
                title="Expected vs Actual Progress by Wave",
                barmode="stack"
            )

            fig_progress.update_layout(
                yaxis_title="Planned Completed Count and Actually Completed Count",
                xaxis_title="Wave",
                yaxis_range=[0, 100]
            )

            st.plotly_chart(fig_progress, use_container_width=True)

# -------------------------
# DRILL-THROUGH COMPONENT DETAIL
# -------------------------
elif page == "Drill-Through Component Detail":
    st.title("Drill-Through – Component Detail")

    if require_upload():
        df = st.session_state.uploaded_data

        if "Wave" in df.columns:
            selected_wave = st.selectbox(
                "Select Wave",
                sorted(df["Wave"].dropna().unique())
            )
            drill_df = df[df["Wave"] == selected_wave]
        else:
            drill_df = df

        display_cols = [
            "Task Name",
            "Wave",
            "Actual Finish",
            "Actual Start",
            "Filing Status"
        ]

        display_cols = [col for col in display_cols if col in drill_df.columns]

        st.subheader("Component-Level Details")
        st.dataframe(
            drill_df[display_cols],
            use_container_width=True,
            hide_index=True
        )

        csv = drill_df[display_cols].to_csv(index=False).encode("utf-8")

        st.download_button(
            "Download Drill-Through Data",
            data=csv,
            file_name="drill_through_component_detail.csv",
            mime="text/csv"
        )

# -------------------------
# NON-ROLLING MODULE VIEW
# -------------------------
elif page == "Non-Rolling Module View":
    st.title("Non-Rolling Submission – Module View")

    if require_upload():
        df = st.session_state.uploaded_data

        if "Module" not in df.columns:
            st.error("The uploaded file does not contain a Module column.")
        else:
            module_summary = get_module_summary(df)

            st.subheader("Module Summary")

            st.dataframe(
                module_summary[
                    [
                        "Module",
                        "Actually_Completed_Count",
                        "Incomplete_Count",
                        "Planned_Completed_Count",
                        "Total_Records",
                        "Completion_Rate"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )

            col1, col2 = st.columns(2)

            with col1:
                fig_module_completion = px.bar(
                    module_summary,
                    x="Module",
                    y="Completion_Rate",
                    title="Completion Rate by Module",
                    text="Completion_Rate"
                )
                fig_module_completion.update_traces(
                    texttemplate="%{text:.2f}%",
                    textposition="outside"
                )
                fig_module_completion.update_layout(
                    yaxis_title="Completion Rate (%)",
                    xaxis_title="Module",
                    yaxis_range=[0, 100]
                )
                st.plotly_chart(fig_module_completion, use_container_width=True)

            with col2:
                fig_module_incomplete = px.bar(
                    module_summary,
                    x="Module",
                    y="Incomplete_Count",
                    title="Incomplete Documents by Module",
                    text="Incomplete_Count"
                )
                fig_module_incomplete.update_layout(
                    yaxis_title="Incomplete Count",
                    xaxis_title="Module"
                )
                st.plotly_chart(fig_module_incomplete, use_container_width=True)

            selected_module = st.selectbox(
                "Select Module for Details",
                sorted(df["Module"].dropna().unique())
            )

            module_detail = df[df["Module"] == selected_module]

            display_cols = [
                "Task Name",
                "Module",
                "Actual Finish",
                "Actual Start",
                "Filing Status"
            ]

            display_cols = [col for col in display_cols if col in module_detail.columns]

            st.subheader("Module-Level Details")
            st.dataframe(
                module_detail[display_cols],
                use_container_width=True,
                hide_index=True
            )

# -------------------------
# ANCHOR DATES PAGE
# -------------------------
elif page == "Anchor Dates":
    st.title("Manual Anchor Dates Entry")
    st.write("Add up to 18 anchor dates.")

    with st.form("anchor_date_form"):
        anchor_name = st.text_input("Anchor Date Name")
        anchor_date = st.date_input("Date")
        anchor_status = st.selectbox(
            "Status",
            ["Complete", "In Progress", "Not Started"]
        )

        submitted = st.form_submit_button("Add Anchor Date")

        if submitted:
            if len(st.session_state.anchor_dates) >= 18:
                st.warning("You can only add up to 18 anchor dates.")
            elif anchor_name.strip() == "":
                st.warning("Anchor Date Name cannot be empty.")
            else:
                new_row = pd.DataFrame(
                    [[anchor_name, anchor_date, anchor_status]],
                    columns=["Anchor Date", "Date", "Status"]
                )

                st.session_state.anchor_dates = pd.concat(
                    [st.session_state.anchor_dates, new_row],
                    ignore_index=True
                )

                st.success("Anchor date added.")

    st.subheader("Current Anchor Dates")

    if not st.session_state.anchor_dates.empty:
        st.dataframe(st.session_state.anchor_dates, use_container_width=True)

        row_to_delete = st.selectbox(
            "Select row to remove",
            options=st.session_state.anchor_dates.index,
            format_func=lambda x: f"Row {x + 1}: {st.session_state.anchor_dates.loc[x, 'Anchor Date']}"
        )

        if st.button("Remove Selected Anchor Date"):
            st.session_state.anchor_dates = (
                st.session_state.anchor_dates
                .drop(row_to_delete)
                .reset_index(drop=True)
            )
            st.success("Anchor date removed.")

        csv = st.session_state.anchor_dates.to_csv(index=False).encode("utf-8")

        st.download_button(
            "Download Anchor Dates as CSV",
            data=csv,
            file_name="anchor_dates.csv",
            mime="text/csv"
        )
    else:
        st.info("No anchor dates added yet.")