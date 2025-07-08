import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel
import pandas as pd
from rapidfuzz import fuzz

selected_mode = None
file_paths = []
column_vars = {}
match_vars = []

fuzzy_threshold = 87
threshold_options = {
    "Complete Match – 100%": 100,
    "Strict – 95%": 95,
    "Balanced – 87%": 87,
    "Loose – 75%": 75,
    "Very Loose – 65%": 65
}

def set_threshold(selection):
    global fuzzy_threshold
    fuzzy_threshold = threshold_options.get(selection, 87)

def normalize(val):
    if pd.isnull(val):
        return ""
    return str(val).strip().lower()

def choose_mode(mode):
    global selected_mode, file_paths
    selected_mode = mode
    label_status.config(text=f"Mode Selected: {mode}")
    btn_select_single.config(state=tk.DISABLED)
    btn_select_multiple.config(state=tk.DISABLED)
    btn_check_duplicates.config(state=tk.DISABLED)
    btn_export_cleaned.config(state=tk.DISABLED)
    btn_export_report.config(state=tk.DISABLED)
    file_paths = []
    column_vars.clear()
    match_vars.clear()
    entry_file_paths.delete(1.0, tk.END)
    selected_columns_display.config(text="")
    if mode == "Single":
        btn_select_single.config(state=tk.NORMAL)
    elif mode == "Multiple":
        btn_select_multiple.config(state=tk.NORMAL)

def update_column_selection(columns):
    column_vars.clear()
    for col in columns:
        column_vars[col] = tk.BooleanVar()

def get_selected_columns():
    return [col for col, var in column_vars.items() if var.get()]

def show_column_selector():
    if not column_vars:
        messagebox.showinfo("Info", "No columns loaded yet.")
        return

    top = Toplevel(root)
    top.title("Select Columns")
    top.geometry("300x400")

    canvas = tk.Canvas(top)
    scrollbar = tk.Scrollbar(top, orient="vertical", command=canvas.yview)
    frame = tk.Frame(canvas)
    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    for col, var in column_vars.items():
        tk.Checkbutton(frame, text=col, variable=var).pack(anchor="w")

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def confirm():
        selected_columns_display.config(text=", ".join(get_selected_columns()))
        top.destroy()

    tk.Button(top, text="OK", command=confirm).pack(pady=5)

def select_single_file():
    global file_paths
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if file_path:
        file_paths = [file_path]
        entry_file_paths.delete(1.0, tk.END)
        entry_file_paths.insert(tk.END, file_path)
        df = pd.read_excel(file_path, nrows=1)
        update_column_selection(df.columns.tolist())
        selected_columns_btn.config(state=tk.NORMAL)
        btn_check_duplicates.config(state=tk.NORMAL)

def select_multiple_files():
    global file_paths
    files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if files:
        for file in files:
            if file not in file_paths:
                file_paths.append(file)
        entry_file_paths.delete(1.0, tk.END)
        entry_file_paths.insert(tk.END, "\n".join(file_paths))
        df = pd.read_excel(file_paths[0], nrows=1)
        update_column_selection(df.columns.tolist())
        selected_columns_btn.config(state=tk.NORMAL)
        btn_check_duplicates.config(state=tk.NORMAL)

def find_fuzzy_matches(data):
    matches = []
    for i in range(len(data)):
        for j in range(i + 1, len(data)):
            val1, file1 = data[i]
            val2, file2 = data[j]
            score = fuzz.ratio(val1, val2)
            if score >= fuzzy_threshold:
                matches.append((val1, file1, val2, file2, score))
    return matches

def display_matches(matches):
    for widget in match_frame.winfo_children():
        widget.destroy()
    match_vars.clear()
    if not matches:
        tk.Label(match_frame, text="No fuzzy matches found.").pack()
        return

    canvas = tk.Canvas(match_frame, height=400)
    scrollbar = tk.Scrollbar(match_frame, orient="vertical", command=canvas.yview)
    inner_frame = tk.Frame(canvas)
    inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    for a, f1, b, f2, score in matches:
        var_a = tk.BooleanVar(value=True)
        var_b = tk.BooleanVar(value=True)
        match_vars.append((var_a, a, f1))
        match_vars.append((var_b, b, f2))
        display_text = f'{a} [{f1}] ⟷ {b} [{f2}]  [Score: {score}]'
        tk.Label(inner_frame, text=display_text, wraplength=750, justify="left", font=("Arial", 9, "bold")).pack(anchor="w", pady=(5, 0))
        tk.Checkbutton(inner_frame, text=f'Remove {a} from {f1}', variable=var_a).pack(anchor="w", padx=15)
        tk.Checkbutton(inner_frame, text=f'Remove {b} from {f2}', variable=var_b).pack(anchor="w", padx=15)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

def compare_excel_files():
    selected_columns = get_selected_columns()
    if not file_paths or not selected_columns:
        messagebox.showerror("Error", "Please select files and columns.")
        return

    try:
        data = []
        for path in file_paths:
            df = pd.read_excel(path)
            fname = path.split("/")[-1] if "/" in path else path.split("\\")[-1]
            for col in selected_columns:
                if col in df.columns:
                    for val in df[col].dropna():
                        data.append((normalize(val), fname))

        matches = find_fuzzy_matches(data)
        display_matches(matches)
        btn_export_cleaned.config(state=tk.NORMAL)
        btn_export_report.config(state=tk.NORMAL)

    except Exception as e:
        messagebox.showerror("Error", str(e))

def export_cleaned():
    selected_columns = get_selected_columns()
    if not file_paths or not selected_columns:
        messagebox.showerror("Error", "Please select files and columns.")
        return

    matched_values_by_file = {}
    for var, val, file in match_vars:
        if var.get():
            key = file.strip()
            if key not in matched_values_by_file:
                matched_values_by_file[key] = set()
            matched_values_by_file[key].add(normalize(val))

    cleaned_data = []
    try:
        for path in file_paths:
            df = pd.read_excel(path)
            fname = path.split("/")[-1] if "/" in path else path.split("\\")[-1]
            remove_vals = matched_values_by_file.get(fname, set())

            to_remove_indices = set()
            for col in selected_columns:
                if col in df.columns:
                    for idx, val in df[col].items():
                        if normalize(val) in remove_vals:
                            to_remove_indices.add(idx)

            df_cleaned = df.drop(index=to_remove_indices)
            cleaned_data.append(df_cleaned)

        final = pd.concat(cleaned_data, ignore_index=True)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if save_path:
            final.to_excel(save_path, index=False)
            messagebox.showinfo("Success", "Cleaned file exported successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to clean and export: {e}")

def export_report():
    if not match_vars:
        messagebox.showerror("Error", "No match data to export!")
        return

    matches_to_export = []
    seen_pairs = set()
    for i in range(0, len(match_vars), 2):
        var_a, val_a, file_a = match_vars[i]
        var_b, val_b, file_b = match_vars[i + 1]
        key = tuple(sorted((normalize(val_a), normalize(val_b))))
        if key in seen_pairs:
            continue
        seen_pairs.add(key)
        score = fuzz.ratio(normalize(val_a), normalize(val_b))
        matches_to_export.append({
            "Value A": val_a,
            "File A": file_a,
            "Value B": val_b,
            "File B": file_b,
            "Similarity Score": score
        })

    if matches_to_export:
        df = pd.DataFrame(matches_to_export)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if save_path:
            df.to_excel(save_path, sheet_name="Fuzzy Matches", index=False)
            messagebox.showinfo("Success", "Match report exported successfully!")

# GUI Setup
root = tk.Tk()
root.title("Dupsy v3.1 – F. Günışığı Aydoğan - morevnadev")
try:
    root.iconbitmap("dupsy_icon.ico")
except:
    pass  # Skip if icon file is not found

label_title = tk.Label(root, text="Select Comparison Mode", font=("Arial", 12))
label_title.pack(pady=10)

btn_mode_single = tk.Button(root, text="Compare Single Excel File", command=lambda: choose_mode("Single"))
btn_mode_single.pack(pady=5)

btn_mode_multi = tk.Button(root, text="Compare Multiple Excel Files", command=lambda: choose_mode("Multiple"))
btn_mode_multi.pack(pady=5)

label_status = tk.Label(root, text="No mode selected", font=("Arial", 10, "italic"))
label_status.pack(pady=10)

threshold_label = tk.Label(root, text="Threshold")
threshold_label.pack()
threshold_var = tk.StringVar(value="Balanced – 87%")
threshold_menu = tk.OptionMenu(root, threshold_var, *threshold_options.keys(), command=set_threshold)
threshold_menu.pack(pady=5)

btn_select_single = tk.Button(root, text="Select Single File", command=select_single_file, state=tk.DISABLED)
btn_select_single.pack(pady=5)

btn_select_multiple = tk.Button(root, text="Select Multiple Files", command=select_multiple_files, state=tk.DISABLED)
btn_select_multiple.pack(pady=5)

entry_file_paths = tk.Text(root, height=4, width=80)
entry_file_paths.pack(pady=5)

selected_columns_btn = tk.Button(root, text="Select Columns", command=show_column_selector, state=tk.DISABLED)
selected_columns_btn.pack(pady=5)

selected_columns_display = tk.Label(root, text="", wraplength=700, justify="left")
selected_columns_display.pack()

btn_check_duplicates = tk.Button(root, text="Dupsify", command=compare_excel_files, state=tk.DISABLED)
btn_check_duplicates.pack(pady=10)

btn_export_report = tk.Button(root, text="Export Report", command=export_report, state=tk.DISABLED)
btn_export_report.pack(pady=5)

btn_export_cleaned = tk.Button(root, text="Export Cleaned File", command=export_cleaned, state=tk.DISABLED)
btn_export_cleaned.pack(pady=5)

match_frame = tk.Frame(root, height=400, width=800)
match_frame.pack(pady=10, fill="both", expand=True)

label_footer = tk.Label(root, text="Developed by F. Günışığı Aydoğan", font=("Arial", 8), anchor="e")
label_footer.pack(side="bottom", fill="x", padx=10, pady=5)

root.mainloop()