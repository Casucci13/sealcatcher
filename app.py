# app.py — Seal Catcher (Streamlit Edition)
import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# --- Configuration ---
SOURCE_FILE = "dataPP.xlsx"
BACKUP_FILE = "dataCopy.xlsx"

st.set_page_config(page_title="Seal Catcher", layout="wide")

# --- Helper Functions ---
def load_excel(path):
    try:
        df = pd.read_excel(path)
        st.success(f"Loaded {path}")
        return df
    except Exception as e:
        st.error(f"Failed to load {path}: {e}")
        return pd.DataFrame()

def save_excel(df, path):
    try:
        df.to_excel(path, index=False)
        st.success(f"Saved updates to {path}")
    except Exception as e:
        st.error(f"Failed to save {path}: {e}")

def copy_excel_file(src=SOURCE_FILE, dest=BACKUP_FILE):
    try:
        workbook = load_workbook(src)
        workbook.save(dest)
        st.success(f"Copied {src} → {dest}")
    except Exception as e:
        st.error(f"Failed to copy file: {e}")

# --- UI Layout ---
st.title("📘 Seal Catcher Dashboard")
st.markdown("Manage and back up your data sheets easily.")

st.sidebar.header("Controls")
action = st.sidebar.radio("Choose an action", ["View Data", "Edit & Save", "Copy Workbook"])

if action == "View Data":
    st.subheader("Current Data")
    data = load_excel(SOURCE_FILE)
    st.dataframe(data, use_container_width=True)

elif action == "Edit & Save":
    st.subheader("Edit and Save Changes")
    data = load_excel(SOURCE_FILE)
    edited_data = st.data_editor(data, use_container_width=True, key="editor")
    if st.button("💾 Save Updates"):
        save_excel(edited_data, SOURCE_FILE)

elif action == "Copy Workbook":
    st.subheader("Copy Workbook")
    if st.button("📑 Copy dataPP → dataCopy"):
        copy_excel_file(SOURCE_FILE, BACKUP_FILE)

st.markdown("---")
st.caption("© 2025 Seal Catcher — Streamlit Edition")
