# ─────────────────────────────────────────────
#  fileops.py  —  Save / Load / Export
# ─────────────────────────────────────────────

import json
import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from urllib.parse import urlparse, parse_qs

from openpyxl import Workbook

import state
from ui.spreadsheet import update_spreadsheet_view
from ui.mapping import rebuild_mapping_ui


# ─────────────────────────────────────────────
#  Project Save / Load
# ─────────────────────────────────────────────

def save_project():
    if not state.excel_data and not state.url_var.get():
        messagebox.showwarning("Warning", "Nothing to save!")
        return

    if not state.current_project_path:
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("Project Files", "*.json")]
        )
        if not path:
            return
        state.current_project_path = path

    current_mapping = {eid: var.get() for eid, var in state.mapping_vars.items()}
    project_state = {
        "excel_path":    state.excel_path_var.get(),
        "form_url":      state.url_var.get(),
        "excel_headers": state.excel_headers,
        "excel_data":    state.excel_data,
        "mapping":       current_mapping,
        "entry_ids":     state.entry_ids,
        "num_sections":  state.num_sections
    }

    try:
        with open(state.current_project_path, 'w') as f:
            json.dump(project_state, f, indent=4)

        project_name = os.path.basename(state.current_project_path)
        state.root.title(f"Google Form URL Generator - {project_name}")
        state.preview_text.config(state="normal")
        state.preview_text.insert("1.0", f"SYSTEM: Project saved to {project_name}\n")
        state.preview_text.config(state="disabled")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save: {e}")


def save_project_as():
    old_path = state.current_project_path
    state.current_project_path = None
    save_project()
    if not state.current_project_path:
        state.current_project_path = old_path


def load_project():
    file_path = filedialog.askopenfilename(filetypes=[("Project JSON", "*.json")])
    if not file_path:
        return

    try:
        with open(file_path, 'r') as f:
            data = json.load(f)

        state.excel_path_var.set(data.get("excel_path", ""))
        state.url_var.set(data.get("form_url", ""))
        state.excel_data    = data.get("excel_data", [])
        state.excel_headers = data.get("excel_headers", [])
        state.entry_ids     = data.get("entry_ids", [])
        state.columns       = state.excel_headers + ["NULL"]

        state.num_sections = data.get("num_sections", 1)
        state.parsed_url = urlparse(state.url_var.get())
        state.query_dict = parse_qs(state.parsed_url.query)
        
        page_history = ",".join(str(i) for i in range(state.num_sections))
        state.query_dict["pageHistory"] = [page_history]
        state.parsed_url = state.parsed_url._replace(
            path=state.parsed_url.path.replace("viewform", "formResponse")
        )

        state.mapping_vars.clear()
        saved_mapping = data.get("mapping", {})
        for eid in state.entry_ids:
            var = tk.StringVar(value=saved_mapping.get(eid, "NULL"))
            state.mapping_vars[eid] = var

        rebuild_mapping_ui()
        update_spreadsheet_view()
        state.current_project_path = file_path
        project_name = os.path.basename(file_path)
        state.root.title(f"Google Form URL Generator - {project_name}")
        messagebox.showinfo("Success", "Project Loaded Successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load project: {e}")


# ─────────────────────────────────────────────
#  Export
# ─────────────────────────────────────────────

def export_as_xlsx():
    if not state.excel_data or not state.final_output_urls:
        messagebox.showwarning("Peringatan", "Generate URLs dulu!")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not file_path:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Generated Responses"
    ws.append(state.excel_headers + ["Generated_Form_URL"])
    for idx, row_dict in enumerate(state.excel_data):
        row_values = [row_dict.get(h, "") for h in state.excel_headers]
        url = state.final_output_urls[idx] if idx < len(state.final_output_urls) else ""
        ws.append(row_values + [url])
    wb.save(file_path)
    messagebox.showinfo("Export Successful", f"Saved to {file_path}")


def export_as_csv():
    if not state.excel_data or not state.final_output_urls:
        messagebox.showwarning("Peringatan", "Generate URLs dulu!")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV Files", "*.csv")]
    )
    if not file_path:
        return

    with open(file_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=state.excel_headers + ["Generated_URL"])
        writer.writeheader()
        for idx, row in enumerate(state.excel_data):
            url = state.final_output_urls[idx] if idx < len(state.final_output_urls) else ""
            export_row = row.copy()
            export_row["Generated_URL"] = url
            writer.writerow(export_row)
    messagebox.showinfo("Export", f"CSV saved to {file_path}")


def export_as_txt():
    if not state.final_output_urls:
        messagebox.showwarning("Peringatan", "Generate URLs dulu!")
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )
    if not path:
        return

    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(state.final_output_urls))
        messagebox.showinfo("Sukses", f"Full URLs ({len(state.final_output_urls)} rows) saved!")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal menyimpan file: {e}")
