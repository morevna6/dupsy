import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from io import BytesIO

# --- App State ---
if 'file_paths' not in st.session_state:
    st.session_state.file_paths = []

if 'column_vars' not in st.session_state:
    st.session_state.column_vars = {}

if 'matches' not in st.session_state:
    st.session_state.matches = []

# --- Fuzzy threshold ---
threshold_options = {
    "Complete Match – 100%": 100,
    "Strict – 95%": 95,
    "Balanced – 87%": 87,
    "Loose – 75%": 75,
    "Very Loose – 65%": 65
}

fuzzy_threshold = st.sidebar.selectbox("Select Threshold", list(threshold_options.keys()), index=2)
threshold_value = threshold_options[fuzzy_threshold]

# --- Normalize function ---
def normalize(val):
    if pd.isnull(val):
        return ""
    return str(val).strip().lower()

# --- File uploader ---
mode = st.radio("Select Comparison Mode", ["Single File", "Multiple Files"])

uploaded_files = st.file_uploader(
    "Upload Excel file(s)",
    type=['xlsx', 'xls'],
    accept_multiple_files=(mode == "Multiple Files")
)

# Make uploaded_files always a list
if uploaded_files:
    if not isinstance(uploaded_files, list):
        uploaded_files = [uploaded_files]

    st.session_state.file_paths = uploaded_files

    # Load first file to extract columns
    uploaded_files[0].seek(0)
    df = pd.read_excel(BytesIO(uploaded_files[0].read()))
    columns = df.columns.tolist()
    st.session_state.column_vars = {col: True for col in columns}  # default all selected

# --- Column selection ---
if st.session_state.column_vars:
    selected_columns = st.multiselect(
        "Select Columns to Compare",
        options=list(st.session_state.column_vars.keys()),
        default=[col for col, selected in st.session_state.column_vars.items() if selected]
    )
else:
    selected_columns = []

# --- Fuzzy matching ---
def find_fuzzy_matches(data, threshold):
    matches = []
    for i in range(len(data)):
        for j in range(i + 1, len(data)):
            val1, file1 = data[i]
            val2, file2 = data[j]
            score = fuzz.ratio(val1, val2)
            if score >= threshold:
                matches.append((val1, file1, val2, file2, score))
    return matches

# --- Compare button ---
if st.button("Dupsify") and uploaded_files and selected_columns:
    data = []
    for f in uploaded_files:
        f.seek(0)
        df = pd.read_excel(BytesIO(f.read()))
        fname = f.name
        for col in selected_columns:
            if col in df.columns:
                for val in df[col].dropna():
                    data.append((normalize(val), fname))
    st.session_state.matches = find_fuzzy_matches(data, threshold_value)
    if st.session_state.matches:
        st.success(f"Found {len(st.session_state.matches)} fuzzy matches!")
    else:
        st.info("No fuzzy matches found.")

# --- Display matches and allow removal ---
matches_to_remove = {}
if st.session_state.matches:
    st.subheader("Fuzzy Matches")
    for idx, (val1, file1, val2, file2, score) in enumerate(st.session_state.matches):
        key1 = f"{val1}_{file1}_{idx}_a"
        key2 = f"{val2}_{file2}_{idx}_b"
        col1, col2 = st.columns([3, 3])
        with col1:
            remove1 = st.checkbox(f"Remove {val1} from {file1}", key=key1)
        with col2:
            remove2 = st.checkbox(f"Remove {val2} from {file2}", key=key2)
        matches_to_remove[key1] = remove1
        matches_to_remove[key2] = remove2

# --- Export cleaned ---
def export_cleaned_file():
    matched_values_by_file = {}
    for idx, (val1, file1, val2, file2, score) in enumerate(st.session_state.matches):
        if matches_to_remove.get(f"{val1}_{file1}_{idx}_a"):
            matched_values_by_file.setdefault(file1, set()).add(val1)
        if matches_to_remove.get(f"{val2}_{file2}_{idx}_b"):
            matched_values_by_file.setdefault(file2, set()).add(val2)

    cleaned_data = []
    for f in st.session_state.file_paths:
        f.seek(0)
        df = pd.read_excel(BytesIO(f.read()))
        fname = f.name
        remove_vals = matched_values_by_file.get(fname, set())
        to_remove_indices = set()
        for col in selected_columns:
            if col in df.columns:
                for idx_row, val in df[col].items():
                    if normalize(val) in remove_vals:
                        to_remove_indices.add(idx_row)
        df_cleaned = df.drop(index=to_remove_indices)
        cleaned_data.append(df_cleaned)

    final = pd.concat(cleaned_data, ignore_index=True)
    buffer = BytesIO()
    final.to_excel(buffer, index=False)
    buffer.seek(0)
    st.download_button("Download Cleaned File", buffer, file_name="cleaned.xlsx")

if st.session_state.matches:
    export_cleaned_file()

# --- Export report ---
def export_report_file():
    report_data = []
    seen_pairs = set()
    for idx, (val1, file1, val2, file2, score) in enumerate(st.session_state.matches):
        key = tuple(sorted((normalize(val1), normalize(val2))))
        if key in seen_pairs:
            continue
        seen_pairs.add(key)
        report_data.append({
            "Value A": val1,
            "File A": file1,
            "Value B": val2,
            "File B": file2,
            "Similarity Score": score
        })
    if report_data:
        df_report = pd.DataFrame(report_data)
        buffer = BytesIO()
        df_report.to_excel(buffer, index=False)
        buffer.seek(0)
        st.download_button("Download Match Report", buffer, file_name="match_report.xlsx")

if st.session_state.matches:
    export_report_file()

st.markdown("---")
st.caption("Developed by F. Günışığı Aydoğan")
