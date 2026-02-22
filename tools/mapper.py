# ─────────────────────────────────────────────
#  tools/mapper.py  —  Likert Scale Mapper
# ─────────────────────────────────────────────

import tkinter as tk
from tkinter import ttk, messagebox

import state
from core import save_undo_state, generate_urls
from ui.mapping import build_shuttle_frame
from ui.spreadsheet import update_spreadsheet_view


def map_likert_gui():
    if not state.excel_data or not state.mapping_vars:
        messagebox.showerror("Error", "Load Excel dan URL dulu!")
        return

    map_window = tk.Toplevel(state.root)
    map_window.title("Likert Mapping Tool (Numbers to Text)")
    map_window.geometry("950x700")
    tk.Label(map_window,
             text="Convert randomized numbers or Excel values into text labels.",
             font=("Arial", 9, "italic")).pack(pady=5)

    list_available, list_selected = build_shuttle_frame(
        map_window,
        label_left="Available (EntryID=Original : Map)",
        label_right="Selected for Mapping",
        label_color="purple"
    )
    for eid, var in state.mapping_vars.items():
        original_answer = state.query_dict.get(eid, [""])[0]
        list_available.insert(tk.END, f"{eid}={original_answer} : {var.get()}")

    config_frame = tk.LabelFrame(map_window, text=" Mapping Configuration ", padx=10, pady=10)
    config_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(config_frame, text="Scale Count:").grid(row=0, column=0)
    num_scale = tk.IntVar(value=7)
    tk.Entry(config_frame, textvariable=num_scale, width=5).grid(row=0, column=1, padx=5)

    tk.Label(config_frame, text="Start Index (Usually 1):").grid(row=0, column=2, padx=10)
    start_idx = tk.IntVar(value=1)
    tk.Entry(config_frame, textvariable=start_idx, width=5).grid(row=0, column=3, padx=5)

    tk.Label(config_frame, text="Labels (Comma separated):").grid(row=1, column=0, pady=10)
    labels_var = tk.StringVar(
        value="Sangat Tidak Setuju, Tidak Setuju, Agak Tidak Setuju, "
              "Netral, Agak Setuju, Setuju, Sangat Setuju"
    )
    tk.Entry(config_frame, textvariable=labels_var, width=80).grid(
        row=1, column=1, columnspan=4, padx=5)

    def apply_mapping():
        save_undo_state()
        selected_items = list_selected.get(0, tk.END)
        target_eids    = [item.split("=")[0].strip() for item in selected_items]
        labels         = [l.strip() for l in labels_var.get().split(",")]
        base_idx       = start_idx.get()

        if not target_eids:
            messagebox.showerror("Error", "Pilih entry untuk di-mapping!")
            return
        if len(labels) != num_scale.get():
            messagebox.showerror(
                "Error",
                f"Jumlah label ({len(labels)}) tidak sesuai dengan "
                f"Scale Count ({num_scale.get()})!"
            )
            return

        count = 0
        for row in state.excel_data:
            for eid in target_eids:
                col = state.mapping_vars[eid].get()
                if col == "NULL" or col not in state.excel_headers:
                    continue
                try:
                    raw = row.get(col)
                    if raw is not None:
                        label_index = int(raw) - base_idx
                        if 0 <= label_index < len(labels):
                            row[col] = labels[label_index]
                            count += 1
                except (ValueError, TypeError):
                    continue

        update_spreadsheet_view()
        try:
            generate_urls()
        except Exception:
            pass
        map_window.destroy()
        messagebox.showinfo("Success", f"Successfully mapped {count} values to text labels!")

    tk.Button(map_window, text="CONVERT NUMBERS TO TEXT LABELS",
              command=apply_mapping, bg="purple", fg="white",
              font=("Arial", 10, "bold"), height=2).pack(fill="x", padx=10, pady=10)
