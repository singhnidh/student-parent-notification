
import streamlit as st
import pandas as pd
import re

from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)


# =========================================================
# 1. PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="SET | Student Academic Portal",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# 2. MODERN FRONTEND DESIGN
# =========================================================

st.markdown("""
<style>

/* Application background */
.stApp {
    background: #f5f7fc;
    font-family: 'Segoe UI', sans-serif;
}

/* Main content */
.block-container {
    max-width: 1500px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}

/* Main headings */
h1, h2, h3 {
    color: #172554;
    font-weight: 750;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #172554;
    border-right: 1px solid #243b75;
}

section[data-testid="stSidebar"] * {
    color: #ffffff;
}

section[data-testid="stSidebar"] div[data-testid="stRadio"] label {
    padding: 9px 12px;
    border-radius: 9px;
}

/* Metric cards */
div[data-testid="stMetric"] {
    background: #ffffff;
    padding: 22px;
    border-radius: 15px;
    border: 1px solid #e5eaf3;
    box-shadow: 0 4px 15px rgba(30, 50, 90, 0.05);
    transition: 0.2s ease;
}

div[data-testid="stMetric"]:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 22px rgba(30, 50, 90, 0.10);
}

div[data-testid="stMetricLabel"] {
    color: #64748b;
    font-weight: 600;
}

div[data-testid="stMetricValue"] {
    color: #172554;
    font-weight: 800;
}

/* Buttons */
.stButton > button,
.stDownloadButton > button {
    background: linear-gradient(120deg, #304da0, #7050c8);
    color: white !important;
    border: none;
    border-radius: 10px;
    padding: 10px 20px;
    font-weight: 650;
    min-height: 43px;
    transition: 0.2s ease;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background: linear-gradient(120deg, #243d83, #5939ae);
    box-shadow: 0 5px 15px rgba(70, 70, 170, 0.20);
    transform: translateY(-2px);
}

/* File uploader */
div[data-testid="stFileUploader"] {
    background: #ffffff;
    border: 2px dashed #a8b8e5;
    border-radius: 14px;
    padding: 15px;
}

/* Dataframes */
div[data-testid="stDataFrame"] {
    border: 1px solid #e3e8f1;
    border-radius: 12px;
    overflow: hidden;
}

/* Input controls */
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div {
    border-radius: 9px;
    border-color: #d4dbea;
}

/* Text area */
.stTextArea textarea {
    border-radius: 10px;
    background: white;
    border: 1px solid #dce3ef;
}

/* Expanders */
div[data-testid="stExpander"] {
    background: white;
    border: 1px solid #e5eaf3;
    border-radius: 12px;
}

/* Alerts */
div[data-testid="stAlert"] {
    border-radius: 10px;
}

/* Tabs */
button[data-baseweb="tab"] {
    font-weight: 650;
}

/* Horizontal divider */
hr {
    border-color: #e1e7f0;
}

/* Hide default Streamlit menu/footer */
#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# 3. REQUIRED EXCEL COLUMNS
# =========================================================

REQUIRED_COLUMNS = [
    "USN",
    "NAME",
    "EXAM_TYPE",
    "SUBJECT",
    "MARKS",
    "MAX_MARKS",
    "FATHER_MOBILE",
    "MOTHER_MOBILE",
    "FATHER_WA_CONSENT",
    "MOTHER_WA_CONSENT",
    "FATHER_SMS_CONSENT",
    "MOTHER_SMS_CONSENT"
]


# =========================================================
# 4. PDF GENERATION FUNCTION
# =========================================================

def create_student_pdf(student_df, exam_type):

    buffer = BytesIO()

    student_name = str(student_df["NAME"].iloc[0])
    usn = str(student_df["USN"].iloc[0])

    pdf = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm
    )

    styles = getSampleStyleSheet()

    subject_style = ParagraphStyle(
        "SubjectStyle",
        parent=styles["Normal"],
        fontSize=9,
        leading=12
    )

    content = []

    content.append(
        Paragraph(
            "<b>SCHOOL OF ENGINEERING &amp; TECHNOLOGY</b>",
            styles["Title"]
        )
    )

    content.append(
        Paragraph(
            "S-VYASA Deemed to be University",
            styles["Heading2"]
        )
    )

    content.append(Spacer(1, 12))

    content.append(
        Paragraph(
            "<b>STUDENT MARKS STATEMENT</b>",
            styles["Heading2"]
        )
    )

    content.append(Spacer(1, 12))

    content.append(
        Paragraph(
            f"<b>Student Name:</b> {escape(student_name)}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"<b>USN:</b> {escape(usn)}",
            styles["Normal"]
        )
    )

    content.append(
        Paragraph(
            f"<b>Examination:</b> {escape(str(exam_type))}",
            styles["Normal"]
        )
    )

    content.append(Spacer(1, 18))

    table_data = [
        [
            "Sl. No.",
            "Subject",
            "Marks Obtained",
            "Maximum Marks"
        ]
    ]

    for index, (_, row) in enumerate(
        student_df.iterrows(),
        start=1
    ):

        raw_marks = row["MARKS"]

        if pd.isna(raw_marks) or str(raw_marks).strip() == "":
            marks_text = "Not Entered"

        elif str(raw_marks).strip().upper() == "AB":
            marks_text = "AB"

        else:
            marks_text = str(raw_marks).strip()

        max_marks = row["MAX_MARKS"]

        if pd.isna(max_marks):
            max_text = "-"
        else:
            max_text = str(max_marks)

        table_data.append([
            str(index),
            Paragraph(
                escape(str(row["SUBJECT"])),
                subject_style
            ),
            marks_text,
            max_text
        ])

    marks_table = Table(
        table_data,
        colWidths=[
            18 * mm,
            85 * mm,
            35 * mm,
            35 * mm
        ],
        repeatRows=1
    )

    marks_table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#233B79")
            ),
            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white
            ),
            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor("#CBD5E1")
            ),
            (
                "ALIGN",
                (0, 0),
                (0, -1),
                "CENTER"
            ),
            (
                "ALIGN",
                (2, 1),
                (-1, -1),
                "CENTER"
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),
            (
                "ROWBACKGROUNDS",
                (0, 1),
                (-1, -1),
                [
                    colors.white,
                    colors.HexColor("#F0F4FA")
                ]
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                9
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                9
            )
        ])
    )

    content.append(marks_table)
    content.append(Spacer(1, 18))

    marks_valid = (
        student_df["MARKS_NUMERIC"].notna().all()
        and student_df["MAX_MARKS_NUMERIC"].notna().all()
    )

    if marks_valid:

        total = student_df["MARKS_NUMERIC"].sum()
        maximum = student_df["MAX_MARKS_NUMERIC"].sum()

        content.append(
            Paragraph(
                f"<b>Total Marks:</b> {total:g} / {maximum:g}",
                styles["Normal"]
            )
        )

        if maximum > 0:

            percentage = (total / maximum) * 100

            content.append(
                Paragraph(
                    f"<b>Percentage:</b> {percentage:.2f}%",
                    styles["Normal"]
                )
            )

    else:

        content.append(
            Paragraph(
                "<b>Result:</b> Incomplete marks. "
                "Percentage has not been calculated.",
                styles["Normal"]
            )
        )

    content.append(Spacer(1, 30))

    content.append(
        Paragraph(
            "Generated electronically by the "
            "Student Marks Management System.",
            styles["Italic"]
        )
    )

    pdf.build(content)

    buffer.seek(0)

    return buffer.getvalue()


# =========================================================
# 5. HELPER FUNCTIONS
# =========================================================

def safe_filename(value):

    return re.sub(
        r"[^A-Za-z0-9_-]",
        "_",
        str(value)
    )


def format_marks(value):

    if pd.isna(value) or str(value).strip() == "":
        return "Not Entered"

    if str(value).strip().upper() == "AB":
        return "Absent (AB)"

    return str(value).strip()


def get_student_status(student_df):

    raw_marks = student_df["MARKS"].fillna("").astype(str).str.strip()

    if raw_marks.eq("").any():
        return "Marks Pending"

    if raw_marks.str.upper().eq("AB").any():
        return "Absent in Subject"

    if student_df["MARKS_NUMERIC"].isna().any():
        return "Check Marks"

    return "Complete"


# =========================================================
# 6. SIDEBAR AND NAVIGATION
# =========================================================

with st.sidebar:

    st.markdown("""
    <div style="text-align:center; padding:18px 5px 12px;">
        <div style="font-size:42px;">🎓</div>
        <div style="
            font-size:21px;
            font-weight:800;
            letter-spacing:1px;
        ">
            SET PORTAL
        </div>
        <div style="font-size:12px; opacity:0.75;">
            Academic Management System
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    st.markdown("### MAIN MENU")

    page = st.radio(
        "Navigation",
        [
            "Dashboard",
            "Upload Excel",
            "Student Records",
            "Performance Analytics",
            "PDF Reports",
            "Parent Notifications"
        ],
        label_visibility="collapsed"
    )

    st.divider()

    st.markdown("""
    <div style="font-size:12px; line-height:1.8;">
        <b>S-VYASA</b><br>
        School of Engineering & Technology<br>
        Academic Performance Portal
    </div>
    """, unsafe_allow_html=True)


# =========================================================
# 7. APPLICATION HEADER
# =========================================================

st.markdown("""
<div style="
    background:linear-gradient(115deg,#172554,#304b98,#7050c8);
    padding:30px 34px;
    border-radius:18px;
    margin-bottom:25px;
    box-shadow:0 8px 25px rgba(30,45,100,0.14);
">

<div style="
        color:#c7d2fe;
        font-size:12px;
        font-weight:750;
        letter-spacing:2px;
        margin-bottom:10px;
    ">
        S-VYASA DEEMED TO BE UNIVERSITY
</div>

<div style="
        color:white;
        font-size:29px;
        font-weight:800;
        line-height:1.3;
        margin-bottom:8px;
    ">
        Student Marks & Parent Notification System
</div>

<div style="
        color:#e0e7ff;
        font-size:14px;
    ">
        School of Engineering & Technology |
        Academic Performance Management
    </div>

</div>
""", unsafe_allow_html=True)


# =========================================================
# 8. GLOBAL EXCEL UPLOAD
# =========================================================

st.markdown("""
<div style="
    color:#64748b;
    font-size:13px;
    font-weight:650;
    margin-bottom:7px;
">
    CURRENT EXAMINATION DATA
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload student Excel workbook",
    type=["xlsx"],
    key="main_excel_upload"
)


# =========================================================
# 9. LOAD AND VALIDATE EXCEL DATA
# =========================================================

df = None

if uploaded_file is not None:

    try:

        df = pd.read_excel(
            uploaded_file,
            dtype={"USN": str}
        )

        df.columns = df.columns.str.strip()

        missing_columns = [
            col for col in REQUIRED_COLUMNS
            if col not in df.columns
        ]

        if missing_columns:

            st.error(
                "Missing Excel columns: "
                + ", ".join(missing_columns)
            )

            df = None

        else:

            text_columns = [
                "USN",
                "NAME",
                "EXAM_TYPE",
                "SUBJECT"
            ]

            for col in text_columns:

                df[col] = (
                    df[col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

            df["MARKS"] = df["MARKS"].apply(
                lambda x: x.strip()
                if isinstance(x, str)
                else x
            )

            df["MARKS_NUMERIC"] = pd.to_numeric(
                df["MARKS"],
                errors="coerce"
            )

            df["MAX_MARKS_NUMERIC"] = pd.to_numeric(
                df["MAX_MARKS"],
                errors="coerce"
            )

            if df.empty:
                st.warning("The uploaded Excel file is empty.")
                df = None

            else:
                st.success("Excel file uploaded and validated successfully.")

    except Exception as e:

        st.error(f"Unable to read Excel file: {e}")
        df = None


# =========================================================
# 10. NO DATA SCREEN
# =========================================================

if df is None:

    st.markdown("""
    <div style="
        background:white;
        border:1px solid #e3e8f2;
        border-radius:18px;
        padding:40px;
        text-align:center;
        margin-top:20px;
    ">
        <div style="font-size:52px;">📂</div>
        <div style="
            font-size:23px;
            font-weight:800;
            color:#172554;
            margin:12px 0;
        ">
            Welcome to the Academic Portal
        </div>
        <div style="
            color:#64748b;
            font-size:14px;
            max-width:550px;
            margin:auto;
        ">
            Upload your student marks Excel workbook above
            to access the dashboard, student records,
            performance analytics and PDF reports.
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.stop()


# =========================================================
# 11. COMMON EXAMINATION DATA
# =========================================================

exams = sorted(
    df.loc[
        df["EXAM_TYPE"] != "",
        "EXAM_TYPE"
    ].unique().tolist()
)

if not exams:

    st.warning("No examination names found in the Excel file.")
    st.stop()


# Keep the selected exam in session state
if (
    "selected_exam" not in st.session_state
    or st.session_state["selected_exam"] not in exams
):
    st.session_state["selected_exam"] = exams[0]


# =========================================================
# 12. DASHBOARD PAGE
# =========================================================

if page == "Dashboard":

    st.markdown("## Academic Overview")

    st.caption(
        "A consolidated view of the uploaded student examination data."
    )

    selected_exam = st.selectbox(
        "Select Examination for Dashboard",
        exams,
        index=exams.index(st.session_state["selected_exam"])
    )

    st.session_state["selected_exam"] = selected_exam

    exam_df = df[
        df["EXAM_TYPE"] == selected_exam
    ].copy()

    total_students = exam_df["USN"].nunique()
    total_subjects = exam_df["SUBJECT"].nunique()
    total_records = len(exam_df)

    complete_records = exam_df[
        exam_df["MARKS_NUMERIC"].notna()
        & exam_df["MAX_MARKS_NUMERIC"].notna()
    ]

    pending_records = exam_df[
        exam_df["MARKS"].isna()
        | exam_df["MARKS"].astype(str).str.strip().eq("")
    ]

    absent_records = exam_df[
        exam_df["MARKS"].astype(str).str.strip().str.upper() == "AB"
    ]

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total Students", total_students)
    col2.metric("Total Subjects", total_subjects)
    col3.metric("Complete Marks", len(complete_records))
    col4.metric("Pending Marks", len(pending_records))

    st.markdown("### Examination Summary")

    col5, col6, col7 = st.columns(3)

    col5.metric("Total Records", total_records)
    col6.metric("Absent Records", len(absent_records))
    col7.metric("Examinations", len(exams))

    st.divider()

    left, right = st.columns([1.15, 0.85])

    with left:

        st.markdown("### Subject-wise Average Marks")

        subject_average = (
            exam_df
            .groupby("SUBJECT")["MARKS_NUMERIC"]
            .mean()
            .dropna()
            .sort_values(ascending=False)
        )

        if not subject_average.empty:

            st.bar_chart(
                subject_average,
                color="#5269bd"
            )

        else:

            st.info("No numeric marks available for this examination.")

    with right:

        st.markdown("### Marks Entry Status")

        status_data = pd.DataFrame({
            "Status": [
                "Numeric marks entered",
                "Absent",
                "Blank / Pending"
            ],
            "Records": [
                len(complete_records),
                len(absent_records),
                len(pending_records)
            ]
        })

        st.dataframe(
            status_data,
            hide_index=True,
            use_container_width=True
        )

        st.caption(
            "Status counts are based on subject-level records."
        )

    st.divider()

    st.markdown("### Recent Student Records")

    recent_students = (
        exam_df[
            ["USN", "NAME", "SUBJECT", "MARKS"]
        ]
        .head(10)
        .copy()
    )

    recent_students.columns = [
        "USN",
        "Student Name",
        "Subject",
        "Marks"
    ]

    st.dataframe(
        recent_students,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# 13. UPLOAD EXCEL PAGE
# =========================================================

elif page == "Upload Excel":

    st.markdown("## Excel Data Management")

    st.caption(
        "Upload and verify student examination data."
    )

    st.markdown("""
    <div style="
        background:white;
        border:1px solid #e3e8f2;
        border-radius:14px;
        padding:20px;
        margin:12px 0;
    ">
        <div style="font-size:18px;font-weight:750;color:#172554;">
            Excel Workbook Requirements
        </div>
        <div style="font-size:13px;color:#64748b;margin-top:8px;">
            Your workbook must contain USN, NAME, EXAM_TYPE,
            SUBJECT, MARKS, MAX_MARKS, parent mobile numbers
            and parent consent columns.
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Uploaded Data Preview")

    st.write(
        f"**Total rows:** {len(df)}"
    )

    st.write(
        f"**Total students:** {df['USN'].nunique()}"
    )

    st.write(
        f"**Total examinations:** {df['EXAM_TYPE'].nunique()}"
    )

    with st.expander("View Complete Uploaded Data"):

        st.dataframe(
            df[REQUIRED_COLUMNS],
            use_container_width=True,
            hide_index=True
        )

    st.markdown("### Column Validation")

    validation_data = pd.DataFrame({
        "Required Column": REQUIRED_COLUMNS,
        "Status": [
            "Available" if col in df.columns else "Missing"
            for col in REQUIRED_COLUMNS
        ]
    })

    st.dataframe(
        validation_data,
        use_container_width=True,
        hide_index=True
    )

    st.success("All required Excel columns are available.")



# =========================================================
# 14. STUDENT RECORDS PAGE
# =========================================================

elif page == "Student Records":

    st.markdown("## Student Records")

    st.caption(
        "Search, filter and view detailed student examination performance."
    )

    # -----------------------------------------------------
    # 1. EXAMINATION SELECTION
    # -----------------------------------------------------

    selected_exam = st.selectbox(
        "Select Examination",
        exams,
        index=exams.index(st.session_state["selected_exam"]),
        key="records_exam"
    )

    st.session_state["selected_exam"] = selected_exam

    exam_df = df[
        df["EXAM_TYPE"] == selected_exam
    ].copy()

    # -----------------------------------------------------
    # 2. SEARCH AND FILTERS
    # -----------------------------------------------------

    st.markdown("### Search and Filters")

    col1, col2 = st.columns([1.5, 1])

    with col1:
        search_value = st.text_input(
            "Search Student",
            placeholder="Enter student name or USN..."
        )

    with col2:
        subject_options = sorted(
            exam_df["SUBJECT"].dropna().unique().tolist()
        )

        selected_subject = st.selectbox(
            "Filter by Subject",
            ["All Subjects"] + subject_options,
            key="records_subject"
        )

    # Search by name or USN
    filtered_df = exam_df.copy()

    if search_value.strip():

        filtered_df = filtered_df[
            filtered_df["NAME"].str.contains(
                search_value.strip(),
                case=False,
                na=False
            )
            |
            filtered_df["USN"].str.contains(
                search_value.strip(),
                case=False,
                na=False
            )
        ]

    # Subject filter
    if selected_subject != "All Subjects":

        filtered_df = filtered_df[
            filtered_df["SUBJECT"] == selected_subject
        ]

    # -----------------------------------------------------
    # 3. MARKS STATUS
    # -----------------------------------------------------

    def get_marks_status(row):

        raw = row["MARKS"]

        if pd.isna(raw) or str(raw).strip() == "":
            return "Pending"

        if str(raw).strip().upper() == "AB":
            return "Absent"

        if pd.isna(row["MARKS_NUMERIC"]):
            return "Check Marks"

        return "Entered"

    filtered_df["MARKS_STATUS"] = filtered_df.apply(
        get_marks_status,
        axis=1
    )

    col3, col4 = st.columns(2)

    with col3:
        status_options = [
            "All Status",
            "Entered",
            "Absent",
            "Pending",
            "Check Marks"
        ]

        selected_status = st.selectbox(
            "Filter by Marks Status",
            status_options,
            key="records_status"
        )

    if selected_status != "All Status":

        filtered_df = filtered_df[
            filtered_df["MARKS_STATUS"] == selected_status
        ]

    # -----------------------------------------------------
    # 4. SUMMARY METRICS
    # -----------------------------------------------------

    st.markdown("### Records Overview")

    total_students = filtered_df["USN"].nunique()
    total_records = len(filtered_df)

    entered_count = (
        filtered_df["MARKS_STATUS"] == "Entered"
    ).sum()

    absent_count = (
        filtered_df["MARKS_STATUS"] == "Absent"
    ).sum()

    pending_count = (
        filtered_df["MARKS_STATUS"] == "Pending"
    ).sum()

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric("Students Found", total_students)
    c2.metric("Total Records", total_records)
    c3.metric("Marks Entered", int(entered_count))
    c4.metric("Absent", int(absent_count))
    c5.metric("Pending", int(pending_count))

    st.divider()

    # -----------------------------------------------------
    # 5. STUDENT LIST
    # -----------------------------------------------------

    st.markdown("### Student List")

    student_summary = (
        filtered_df[
            ["USN", "NAME"]
        ]
        .drop_duplicates()
        .sort_values("NAME")
    )

    if student_summary.empty:

        st.warning(
            "No student records found for the selected search and filters."
        )

    else:

        student_summary.columns = [
            "USN",
            "Student Name"
        ]

        st.dataframe(
            student_summary,
            use_container_width=True,
            hide_index=True
        )

        # -------------------------------------------------
        # 6. INDIVIDUAL STUDENT DETAILS
        # -------------------------------------------------

        st.divider()

        st.markdown("### Individual Student Performance")

        student_options = {
    f"{row['Student Name']} ({row['USN']})": row["USN"]
    for _, row in student_summary.iterrows()
}
        selected_student = st.selectbox(
            "Select Student to View Detailed Marks",
            list(student_options.keys()),
            key="records_student"
        )

        selected_usn = student_options[selected_student]

        student_df = exam_df[
            exam_df["USN"] == selected_usn
        ].copy()

        student_name = student_df["NAME"].iloc[0]

        st.markdown("#### Student Information")

        info1, info2, info3 = st.columns(3)

        info1.metric("Student Name", student_name)
        info2.metric("USN", selected_usn)
        info3.metric("Examination", selected_exam)

        # -------------------------------------------------
        # 7. SUBJECT-WISE MARKS
        # -------------------------------------------------

        st.markdown("#### Subject-wise Marks")

        display_df = student_df[
            ["SUBJECT", "MARKS", "MAX_MARKS"]
        ].copy()

        display_df["Marks Status"] = student_df.apply(
            get_marks_status,
            axis=1
        )

        display_df["Marks Obtained"] = student_df[
            "MARKS"
        ].apply(format_marks)

        display_df["Maximum Marks"] = student_df[
            "MAX_MARKS"
        ].apply(
            lambda x: "-" if pd.isna(x) else str(x)
        )

        display_df = display_df[
            [
                "SUBJECT",
                "Marks Obtained",
                "Maximum Marks",
                "Marks Status"
            ]
        ]

        display_df.columns = [
            "Subject",
            "Marks Obtained",
            "Maximum Marks",
            "Status"
        ]

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )

        # -------------------------------------------------
        # 8. TOTAL MARKS AND PERCENTAGE
        # -------------------------------------------------

        st.markdown("#### Academic Performance Summary")

        marks_valid = (
            student_df["MARKS_NUMERIC"].notna().all()
            and student_df["MAX_MARKS_NUMERIC"].notna().all()
        )

        if marks_valid:

            total_marks = student_df[
                "MARKS_NUMERIC"
            ].sum()

            maximum_marks = student_df[
                "MAX_MARKS_NUMERIC"
            ].sum()

            if maximum_marks > 0:

                percentage = (
                    total_marks / maximum_marks
                ) * 100

                m1, m2, m3 = st.columns(3)

                m1.metric(
                    "Total Marks Obtained",
                    f"{total_marks:g}"
                )

                m2.metric(
                    "Maximum Marks",
                    f"{maximum_marks:g}"
                )

                m3.metric(
                    "Overall Percentage",
                    f"{percentage:.2f}%"
                )

            else:

                st.warning(
                    "Maximum marks must be greater than zero "
                    "to calculate the percentage."
                )

        else:

            st.warning(
                "Some subject marks are absent, pending or invalid. "
                "Complete total marks and percentage cannot be calculated."
            )

        # -------------------------------------------------
        # 9. DOWNLOAD OPTIONS
        # -------------------------------------------------

        st.divider()

        st.markdown("### Download Student Records")

        download_df = student_df[
            [
                "USN",
                "NAME",
                "EXAM_TYPE",
                "SUBJECT",
                "MARKS",
                "MAX_MARKS"
            ]
        ].copy()

        st.download_button(
            "Download Selected Student Records (CSV)",
            data=download_df.to_csv(
                index=False
            ).encode("utf-8"),
            file_name=(
                f"{safe_filename(student_name)}_"
                f"{safe_filename(selected_exam)}.csv"
            ),
            mime="text/csv",
            key="download_selected_student"
        )

    # -----------------------------------------------------
    # 10. DOWNLOAD ALL FILTERED RECORDS
    # -----------------------------------------------------

    st.divider()

    st.markdown("### Export Filtered Records")

    export_df = filtered_df[
        [
            "USN",
            "NAME",
            "EXAM_TYPE",
            "SUBJECT",
            "MARKS",
            "MAX_MARKS",
            "MARKS_STATUS"
        ]
    ].copy()

    st.download_button(
        "Download All Filtered Records (CSV)",
        data=export_df.to_csv(
            index=False
        ).encode("utf-8"),
        file_name=(
            f"Filtered_Student_Records_"
            f"{safe_filename(selected_exam)}.csv"
        ),
        mime="text/csv",
        key="download_filtered_records"
    )



# =========================================================
# 15. PERFORMANCE ANALYTICS PAGE
# =========================================================

elif page == "Performance Analytics":

    st.markdown("## Performance Analytics")

    st.caption(
        "Explore subject-wise results, student performance "
        "and examination statistics."
    )

    # -----------------------------------------------------
    # 1. EXAMINATION SELECTION
    # -----------------------------------------------------

    selected_exam = st.selectbox(
        "Select Examination",
        exams,
        index=exams.index(st.session_state["selected_exam"]),
        key="analytics_exam"
    )

    st.session_state["selected_exam"] = selected_exam

    exam_df = df[
        df["EXAM_TYPE"] == selected_exam
    ].copy()

    # -----------------------------------------------------
    # 2. OVERALL EXAMINATION METRICS
    # -----------------------------------------------------

    total_students = exam_df["USN"].nunique()
    total_subjects = exam_df["SUBJECT"].nunique()

    entered_records = exam_df[
        exam_df["MARKS_NUMERIC"].notna()
        & exam_df["MAX_MARKS_NUMERIC"].notna()
    ]

    absent_records = exam_df[
        exam_df["MARKS"].astype(str).str.strip().str.upper() == "AB"
    ]

    pending_records = exam_df[
        exam_df["MARKS"].isna()
        | exam_df["MARKS"].astype(str).str.strip().eq("")
    ]

    st.markdown("### Examination Overview")

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Total Students", total_students)
    c2.metric("Total Subjects", total_subjects)
    c3.metric("Marks Entered", len(entered_records))
    c4.metric("Pending / Absent", len(pending_records) + len(absent_records))

    st.divider()

    # -----------------------------------------------------
    # 3. SUBJECT-WISE PERFORMANCE
    # -----------------------------------------------------

    st.markdown("### Subject-wise Performance")

    numeric_df = exam_df[
        exam_df["MARKS_NUMERIC"].notna()
        & exam_df["MAX_MARKS_NUMERIC"].notna()
        & (exam_df["MAX_MARKS_NUMERIC"] > 0)
    ].copy()

    if numeric_df.empty:

        st.warning(
            "No valid numeric marks are available for this examination."
        )

    else:

        numeric_df["Percentage"] = (
            numeric_df["MARKS_NUMERIC"]
            / numeric_df["MAX_MARKS_NUMERIC"]
        ) * 100

        subject_stats = (
            numeric_df
            .groupby("SUBJECT")
            .agg(
                Average_Marks=("MARKS_NUMERIC", "mean"),
                Highest_Marks=("MARKS_NUMERIC", "max"),
                Lowest_Marks=("MARKS_NUMERIC", "min"),
                Average_Percentage=("Percentage", "mean"),
                Students_Recorded=("USN", "nunique")
            )
            .round(2)
            .reset_index()
        )

        col1, col2 = st.columns(2)

        with col1:

            st.markdown("#### Average Marks by Subject")

            st.bar_chart(
                subject_stats.set_index("SUBJECT")[
                    "Average_Marks"
                ],
                color="#5269bd"
            )

        with col2:

            st.markdown("#### Average Percentage by Subject")

            st.bar_chart(
                subject_stats.set_index("SUBJECT")[
                    "Average_Percentage"
                ],
                color="#7050c8"
            )

        st.markdown("#### Subject Performance Details")

        subject_display = subject_stats.rename(
            columns={
                "SUBJECT": "Subject",
                "Average_Marks": "Average Marks",
                "Highest_Marks": "Highest Marks",
                "Lowest_Marks": "Lowest Marks",
                "Average_Percentage": "Average Percentage (%)",
                "Students_Recorded": "Students Recorded"
            }
        )

        st.dataframe(
            subject_display,
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            "Download Subject-wise Analysis (CSV)",
            data=subject_display.to_csv(
                index=False
            ).encode("utf-8"),
            file_name=(
                f"Subject_Analysis_{safe_filename(selected_exam)}.csv"
            ),
            mime="text/csv",
            key="download_subject_analysis"
        )

    st.divider()

    # -----------------------------------------------------
    # 4. STUDENT-WISE PERFORMANCE
    # -----------------------------------------------------

    st.markdown("### Student-wise Performance")

    student_results = []

    for (usn, name), group in exam_df.groupby(
        ["USN", "NAME"],
        dropna=False
    ):

        marks_complete = (
            group["MARKS_NUMERIC"].notna().all()
            and group["MAX_MARKS_NUMERIC"].notna().all()
        )

        total_marks = group["MARKS_NUMERIC"].sum(
            min_count=len(group)
        )

        maximum_marks = group["MAX_MARKS_NUMERIC"].sum(
            min_count=len(group)
        )

        if (
            marks_complete
            and pd.notna(total_marks)
            and pd.notna(maximum_marks)
            and maximum_marks > 0
        ):

            percentage = (
                total_marks / maximum_marks
            ) * 100

            result_status = "Complete"

        else:

            total_marks = None
            maximum_marks = None
            percentage = None
            result_status = "Incomplete"

        student_results.append({
            "USN": usn,
            "Student Name": name,
            "Total Marks": total_marks,
            "Maximum Marks": maximum_marks,
            "Percentage": percentage,
            "Result Status": result_status
        })

    student_summary = pd.DataFrame(student_results)

    complete_students = student_summary[
        student_summary["Result Status"] == "Complete"
    ].copy()

    incomplete_students = student_summary[
        student_summary["Result Status"] == "Incomplete"
    ].copy()

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Complete Results",
        len(complete_students)
    )

    col2.metric(
        "Incomplete Results",
        len(incomplete_students)
    )

    if not complete_students.empty:

        average_percentage = complete_students[
            "Percentage"
        ].mean()

        col3.metric(
            "Average Student Percentage",
            f"{average_percentage:.2f}%"
        )

    else:

        col3.metric(
            "Average Student Percentage",
            "N/A"
        )

    if not complete_students.empty:

        complete_students = complete_students.sort_values(
            "Percentage",
            ascending=False
        )

        st.markdown("#### Student Percentage Comparison")

        st.bar_chart(
            complete_students.set_index("Student Name")[
                "Percentage"
            ],
            color="#304da0"
        )

        st.markdown("#### Complete Student Results")

        st.dataframe(
            complete_students.round(2),
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "Complete student percentages are not available yet."
        )

    if not incomplete_students.empty:

        with st.expander("View Students with Incomplete Marks"):

            st.dataframe(
                incomplete_students,
                use_container_width=True,
                hide_index=True
            )

    st.divider()

    # -----------------------------------------------------
    # 5. TOP 10 PERFORMERS
    # -----------------------------------------------------

    st.markdown("### Top 10 Performers")

    if not complete_students.empty:

        top_students = complete_students.head(10).copy()

        top_students.insert(
            0,
            "Rank",
            range(1, len(top_students) + 1)
        )

        top_students["Percentage"] = (
            top_students["Percentage"].round(2)
        )

        st.dataframe(
            top_students,
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            "Download Top 10 Performers (CSV)",
            data=top_students.to_csv(
                index=False
            ).encode("utf-8"),
            file_name=(
                f"Top_10_Performers_{safe_filename(selected_exam)}.csv"
            ),
            mime="text/csv",
            key="download_top_students"
        )

    else:

        st.info(
            "Top performers will appear once complete marks are available."
        )

    st.divider()

    # -----------------------------------------------------
    # 6. DOWNLOAD COMPLETE ANALYTICS
    # -----------------------------------------------------

    st.markdown("### Download Complete Performance Report")

    export_summary = student_summary.copy()

    export_summary["Total Marks"] = (
        export_summary["Total Marks"].round(2)
    )

    export_summary["Maximum Marks"] = (
        export_summary["Maximum Marks"].round(2)
    )

    export_summary["Percentage"] = (
        export_summary["Percentage"].round(2)
    )

    st.download_button(
        "Download Student Performance Report (CSV)",
        data=export_summary.to_csv(
            index=False
        ).encode("utf-8"),
        file_name=(
            f"Student_Performance_{safe_filename(selected_exam)}.csv"
        ),
        mime="text/csv",
        key="download_student_analytics"
    )

    st.caption(
        "Note: Student percentages and rankings are calculated "
        "only when all subject marks and maximum marks are available. "
        "No pass/fail classification is assumed."
    )


# =========================================================
# 16. PDF REPORTS PAGE
# =========================================================

elif page == "PDF Reports":

    st.markdown("## Student Report Centre")

    st.caption(
        "Generate individual student marks statements "
        "or bulk PDF reports."
    )

    selected_exam = st.selectbox(
        "Select Examination",
        exams,
        index=exams.index(st.session_state["selected_exam"]),
        key="pdf_exam"
    )

    st.session_state["selected_exam"] = selected_exam

    exam_df = df[
        df["EXAM_TYPE"] == selected_exam
    ].copy()

    students = (
        exam_df[
            ["USN", "NAME"]
        ]
        .drop_duplicates()
        .sort_values("NAME")
    )

    student_options = {
        f"{row.NAME} ({row.USN})": row.USN
        for row in students.itertuples()
    }

    st.metric(
        "Students in Selected Examination",
        len(student_options)
    )

    st.divider()

    st.markdown("### Individual Student PDF")

    selected_student = st.selectbox(
        "Select Student",
        list(student_options.keys()),
        key="pdf_student"
    )

    selected_usn = student_options[selected_student]

    student_df = exam_df[
        exam_df["USN"] == selected_usn
    ].copy()

    marks_display = student_df[
        ["SUBJECT", "MARKS", "MAX_MARKS"]
    ].copy()

    marks_display.columns = [
        "Subject",
        "Marks Obtained",
        "Maximum Marks"
    ]

    st.dataframe(
        marks_display,
        use_container_width=True,
        hide_index=True
    )

    pdf_data = create_student_pdf(
        student_df,
        selected_exam
    )

    st.download_button(
        "Download Individual Student PDF",
        data=pdf_data,
        file_name=(
            f"{safe_filename(selected_student)}_"
            f"{safe_filename(selected_exam)}.pdf"
        ),
        mime="application/pdf"
    )

    st.divider()

    st.markdown("### Bulk PDF Generation")

    st.write(
        "Generate a separate PDF marks statement for every student "
        "in the selected examination."
    )

    if st.button(
        "Generate All Student PDFs",
        key="generate_bulk_pdfs"
    ):

        zip_buffer = BytesIO()

        with ZipFile(
            zip_buffer,
            "w",
            ZIP_DEFLATED
        ) as zip_file:

            for usn in exam_df["USN"].unique():

                one_student = exam_df[
                    exam_df["USN"] == usn
                ].copy()

                pdf_file = create_student_pdf(
                    one_student,
                    selected_exam
                )

                name = str(
                    one_student["NAME"].iloc[0]
                )

                filename = safe_filename(name)
                safe_usn = safe_filename(usn)

                zip_file.writestr(
                    f"{filename}_{safe_usn}_"
                    f"{safe_filename(selected_exam)}.pdf",
                    pdf_file
                )

        st.session_state["all_student_pdfs"] = (
            zip_buffer.getvalue()
        )

        st.session_state["pdf_exam_type"] = selected_exam

        st.success(
            f"Successfully generated reports for "
            f"{exam_df['USN'].nunique()} students."
        )

    if (
        "all_student_pdfs" in st.session_state
        and st.session_state.get("pdf_exam_type") == selected_exam
    ):

        st.download_button(
            "Download All Student Reports (ZIP)",
            data=st.session_state["all_student_pdfs"],
            file_name=(
                f"All_Student_Reports_"
                f"{safe_filename(selected_exam)}.zip"
            ),
            mime="application/zip"
        )


# =========================================================
# 17. PARENT NOTIFICATIONS PAGE
# =========================================================

elif page == "Parent Notifications":

    st.markdown("## Parent Communication Centre")

    st.caption(
        "Prepare student marks notifications for parents."
    )

    st.info(
        "This section currently generates message previews only. "
        "WhatsApp and SMS sending are not enabled."
    )

    selected_exam = st.selectbox(
        "Select Examination",
        exams,
        index=exams.index(st.session_state["selected_exam"]),
        key="notification_exam"
    )

    st.session_state["selected_exam"] = selected_exam

    exam_df = df[
        df["EXAM_TYPE"] == selected_exam
    ].copy()

    students = (
        exam_df[
            ["USN", "NAME"]
        ]
        .drop_duplicates()
        .sort_values("NAME")
    )

    student_options = {
        f"{row.NAME} ({row.USN})": row.USN
        for row in students.itertuples()
    }

    selected_student = st.selectbox(
        "Select Student",
        list(student_options.keys()),
        key="notification_student"
    )

    selected_usn = student_options[selected_student]

    student_df = exam_df[
        exam_df["USN"] == selected_usn
    ].copy()

    st.markdown("### Student Information")

    col1, col2 = st.columns(2)

    col1.metric(
        "Student Name",
        student_df["NAME"].iloc[0]
    )

    col2.metric(
        "Examination",
        selected_exam
    )

    st.markdown("### Marks Summary")

    notification_marks = student_df[
        ["SUBJECT", "MARKS", "MAX_MARKS"]
    ].copy()

    notification_marks.columns = [
        "Subject",
        "Marks Obtained",
        "Maximum Marks"
    ]

    st.dataframe(
        notification_marks,
        use_container_width=True,
        hide_index=True
    )

    message = (
        f"Dear Parent,\n\n"
        f"Marks of your ward "
        f"{student_df['NAME'].iloc[0]} "
        f"(USN: {selected_usn}) "
        f"in {selected_exam}:\n\n"
    )

    for _, row in student_df.iterrows():

        marks_text = format_marks(row["MARKS"])

        message += (
            f"{row['SUBJECT']}: "
            f"{marks_text}/{row['MAX_MARKS']}\n"
        )

    message += (
        "\nRegards,\n"
        "School of Engineering & Technology\n"
        "S-VYASA Deemed to be University"
    )

    st.markdown("### Generated Parent Message")

    st.text_area(
        "Message Preview",
        value=message,
        height=280
    )

    st.download_button(
        "Download Message as Text File",
        data=message,
        file_name=(
            f"Parent_Message_{safe_filename(selected_usn)}.txt"
        ),
        mime="text/plain"
    )

    st.divider()

    st.markdown("### Parent Consent Status")

    consent_data = pd.DataFrame([
        {
            "Parent": "Father",
            "WhatsApp Consent": student_df["FATHER_WA_CONSENT"].iloc[0],
            "SMS Consent": student_df["FATHER_SMS_CONSENT"].iloc[0]
        },
        {
            "Parent": "Mother",
            "WhatsApp Consent": student_df["MOTHER_WA_CONSENT"].iloc[0],
            "SMS Consent": student_df["MOTHER_SMS_CONSENT"].iloc[0]
        }
    ])

    st.dataframe(
        consent_data,
        use_container_width=True,
        hide_index=True
    )

    st.warning(
        "Do not send messages unless the required parent consent "
        "and institutional approvals are in place. "
        "This application does not send messages yet."
    )


 