
import streamlit as st

st.set_page_config(
    page_title="Student Parent Notification",
    page_icon="📚",
    layout="wide"
)

st.title("Student Marks & Parent Notification System")

st.success("Streamlit is working successfully!")

st.header("Student Dashboard")

col1, col2, col3 = st.columns(3)

col1.metric("Total Students", 0)
col2.metric("Total Subjects", 0)
col3.metric("Messages Sent", 0)

st.subheader("Upload Student Marks")

uploaded_file = st.file_uploader(
    "Upload your Excel file",
    type=["xlsx"]
)

if uploaded_file is not None:
    st.write("Excel file uploaded:", uploaded_file.name)
