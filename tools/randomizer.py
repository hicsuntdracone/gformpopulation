# ─────────────────────────────────────────────
#  tools/randomizer.py  —  Mass Likert Randomizer
# ─────────────────────────────────────────────

import tkinter as tk
from tkinter import ttk, messagebox

import state
from core import _generate_population, save_undo_state, generate_urls
from ui.mapping import build_shuttle_frame
from ui.spreadsheet import update_spreadsheet_view


def randomizer_gui():
    if not state.excel_data or not state.mapping_vars:
        messagebox.showerror("Error", "Load Excel dan URL dulu!")
        return

    rand_window = tk.Toplevel(state.root)
    rand_window.title("Mass Likert Randomizer Tool")
    rand_window.geometry("950x700")
    tk.Label(rand_window,
             text="Pilih Entry ID untuk di-random massal.",
             font=("Arial", 9, "italic")).pack(pady=5)

    list_available, list_selected = build_shuttle_frame(
        rand_window,
        label_left="Available (EntryID=Original [Map])",
        label_right="Selected for Random",
        label_color="red"
    )
    for eid, var in state.mapping_vars.items():
        original_val = state.query_dict.get(eid, [""])[0]
        list_available.insert(tk.END, f"{eid}={original_val} [{var.get()}]")

    config_frame = tk.LabelFrame(rand_window, text=" Random Configuration ", padx=10, pady=10)
    config_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(config_frame, text="Min:").grid(row=0, column=0)
    min_val = tk.IntVar(value=1)
    tk.Entry(config_frame, textvariable=min_val, width=8).grid(row=0, column=1, padx=5)

    tk.Label(config_frame, text="Max:").grid(row=0, column=2, padx=10)
    max_val = tk.IntVar(value=7)
    tk.Entry(config_frame, textvariable=max_val, width=8).grid(row=0, column=3, padx=5)

    tk.Label(config_frame, text="Method:").grid(row=0, column=4, padx=10)
    method_var = tk.StringVar(value="Normal (Bell Curve)")
    ttk.Combobox(config_frame, textvariable=method_var, state="readonly",
                 values=["True Random", "Normal (Bell Curve)",
                         "Skewed Positive", "Skewed Negative"]
                 ).grid(row=0, column=5, padx=5)

    def run_randomizer():
        save_undo_state()
        selected_items = list_selected.get(0, tk.END)
        target_eids    = [item.split("=")[0].strip() for item in selected_items]
        if not target_eids:
            messagebox.showwarning("Warning", "Pilih entry untuk di-random!")
            return

        mini = min_val.get()
        maxi = max_val.get()
        meth = method_var.get()
        n    = len(state.excel_data)

        for eid in target_eids:
            col = state.mapping_vars[eid].get()
            if col == "NULL" or col not in state.excel_headers:
                continue
            population = _generate_population(mini, maxi, meth, n)
            for i, row in enumerate(state.excel_data):
                row[col] = population[i]

        update_spreadsheet_view()
        try:
            generate_urls()
        except Exception:
            pass
        rand_window.destroy()
        messagebox.showinfo("Success",
                            f"Randomization complete for {len(target_eids)} entries!")

    tk.Button(rand_window, text="APPLY RANDOM TO SELECTED ENTRIES",
              command=run_randomizer, bg="red", fg="white",
              font=("Arial", 10, "bold"), height=2).pack(fill="x", padx=10, pady=10)
