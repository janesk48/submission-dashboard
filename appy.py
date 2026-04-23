import streamlit as st
import streamlit.components.v1 as components
import pandas as pd

st.set_page_config(page_title="Submission Dashboard", layout="wide")

# -------------------------
# SESSION STATE
# -------------------------
if "anchor_dates" not in st.session_state:
    st.session_state.anchor_dates = pd.DataFrame(
        columns=["Anchor Date", "Date", "Status"]
    )

# -------------------------
# SIDEBAR
# -------------------------
st.sidebar.title("Navigation")
page = st.sidebar.radio(
    "Go to",
    ["Dashboard", "Upload Excel", "Anchor Dates"]
)

# -------------------------
# DASHBOARD PAGE
# -------------------------
if page == "Dashboard":
    st.title("Non-Rolling Submission – Module View")
    st.markdown("Embedded Power BI dashboard")

    power_bi_url = "https://app.powerbi.com/reportEmbed?reportId=008bc16e-ca99-48b9-9ad9-48dfe44bb8c0&autoAuth=true&ctid=17dcb00c-6941-4050-b69e-bd7eb8951712"

    st.info("If the dashboard appears blank, it may require NJIT Power BI sign-in.")

    components.html(
        f"""
        <iframe
            title="Power BI Dashboard"
            width="100%"
            height="850"
            src="{power_bi_url}"
            frameborder="0"
            allowFullScreen="true">
        </iframe>
        """,
        height=850,
    )

# -------------------------
# UPLOAD EXCEL PAGE
# -------------------------
elif page == "Upload Excel":
    st.title("Upload Excel File")
    st.write("Upload the current input Excel file for processing.")

    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=["xlsx", "xls"]
    )

    if uploaded_file is not None:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names

            st.success("File uploaded successfully.")
            st.write("Sheets found:", sheet_names)

            selected_sheet = st.selectbox("Select sheet to preview", sheet_names)
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

            st.subheader("Preview")
            st.dataframe(df, use_container_width=True)

            # Optional: save dataframe in session state
            st.session_state["uploaded_data"] = df

        except Exception as e:
            st.error(f"Error reading Excel file: {e}")

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
            st.session_state.anchor_dates = st.session_state.anchor_dates.drop(row_to_delete).reset_index(drop=True)
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