# ─────────────────────────────────────────────
#  ui/mapping.py  —  Field mapping panel &
#                    shuttle listbox widget
# ─────────────────────────────────────────────

import tkinter as tk
from tkinter import ttk

import state


# ─────────────────────────────────────────────
#  Shuttle Listbox Widget
# ─────────────────────────────────────────────

def build_shuttle_frame(parent, label_left, label_right, label_color):
    """Build a two-list shuttle widget. Returns (list_available, list_selected)."""
    frame = tk.Frame(parent)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    def make_list(container, label, color=None):
        tk.Label(container, text=label, font=("Arial", 10, "bold"),
                 **({"fg": color} if color else {})).pack()
        lb = tk.Listbox(container, selectmode="extended", font=("Consolas", 10))
        lb.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(container, command=lb.yview)
        sb.pack(side="right", fill="y")
        lb.config(yscrollcommand=sb.set)
        return lb

    left_container  = tk.Frame(frame)
    left_container.pack(side="left", fill="both", expand=True)
    list_available  = make_list(left_container, label_left)

    btn_frame = tk.Frame(frame)
    btn_frame.pack(side="left", padx=10)

    right_container = tk.Frame(frame)
    right_container.pack(side="left", fill="both", expand=True)
    list_selected   = make_list(right_container, label_right, label_color)

    def move_right():
        for i in reversed(list_available.curselection()):
            list_selected.insert(tk.END, list_available.get(i))
            list_available.delete(i)

    def move_all_right():
        for item in list_available.get(0, tk.END):
            list_selected.insert(tk.END, item)
        list_available.delete(0, tk.END)

    def move_left():
        for i in reversed(list_selected.curselection()):
            list_available.insert(tk.END, list_selected.get(i))
            list_selected.delete(i)

    tk.Button(btn_frame, text="Add Selected >",  command=move_right,     width=15).pack(pady=5)
    tk.Button(btn_frame, text="Add All >>",      command=move_all_right, width=15).pack(pady=5)
    tk.Button(btn_frame, text="< Remove",        command=move_left,      width=15).pack(pady=5)

    return list_available, list_selected


# ─────────────────────────────────────────────
#  Mapping Panel Builder
# ─────────────────────────────────────────────

def rebuild_mapping_ui():
    for widget in state.scrollable_frame.winfo_children():
        widget.destroy()

    tk.Label(state.scrollable_frame, text="Mapping Entry ID ke Kolom Excel:",
             font=("Arial", 10, "bold")).pack(pady=5)

    for entry_id, var in state.mapping_vars.items():
        row_frame = tk.Frame(state.scrollable_frame)
        row_frame.pack(fill="x", padx=10, pady=2)

        original_answer = state.query_dict.get(entry_id, [""])[0]
        tk.Label(row_frame, text=f"{entry_id}={original_answer}:",
                 width=40, anchor="w", font=("Consolas", 9)).pack(side="left")

        combo = ttk.Combobox(row_frame, textvariable=var, values=state.columns,
                             state="readonly", width=20)
        combo.pack(side="left", padx=5)
        combo.bind("<Enter>", lambda e: _unbind_scroll())
        combo.bind("<Leave>", lambda e: _bind_scroll())

        tk.Button(row_frame, text="🎲",
                  command=lambda v=var, e=entry_id: _quick_rand_dialog(v, e),
                  bg="#f0f0f0", width=3).pack(side="left")


def _bind_scroll(event=None):
    state.root.bind_all("<MouseWheel>", _on_mousewheel)

def _unbind_scroll(event=None):
    state.root.unbind_all("<MouseWheel>")

def _on_mousewheel(event):
    state.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def _quick_rand_dialog(target_var, eid):
    if not state.excel_data:
        return

    from tkinter import messagebox
    from core import _random_value, save_undo_state
    from ui.spreadsheet import update_spreadsheet_view
    from core import generate_urls

    dialog = tk.Toplevel(state.root)
    dialog.title(f"Quick Random: {eid}")
    dialog.geometry("320x280")
    dialog.transient(state.root)
    dialog.grab_set()

    tk.Label(dialog, text="Min Value:").grid(row=0, column=0, padx=10, pady=10)
    min_ent = tk.Entry(dialog, width=15)
    min_ent.insert(0, "1")
    min_ent.grid(row=0, column=1)

    tk.Label(dialog, text="Max Value:").grid(row=1, column=0, padx=10, pady=10)
    max_ent = tk.Entry(dialog, width=15)
    max_ent.insert(0, "7")
    max_ent.grid(row=1, column=1)

    tk.Label(dialog, text="Method:").grid(row=2, column=0, padx=10, pady=10)
    method_cmb = ttk.Combobox(
        dialog, state="readonly",
        values=["True Random", "Normal (Bell Curve)", "Skewed Positive", "Skewed Negative"]
    )
    method_cmb.set("Normal (Bell Curve)")
    method_cmb.grid(row=2, column=1)

    def confirm():
        save_undo_state()
        try:
            mini = int(min_ent.get())
            maxi = int(max_ent.get())
            meth = method_cmb.get()
            col  = target_var.get()
            if col == "NULL" or col not in state.excel_headers:
                messagebox.showerror("Error",
                    "Pilih kolom Excel yang valid di mapping terlebih dahulu!")
                return
            for row in state.excel_data:
                row[col] = _random_value(mini, maxi, meth)
            update_spreadsheet_view()
            try:
                generate_urls()
            except Exception:
                pass
            dialog.destroy()
        except Exception:
            messagebox.showerror("Error", "Input tidak valid!")

    tk.Button(dialog, text="Apply Random", command=confirm,
              bg="green", fg="white", width=20).grid(row=3, columnspan=2, pady=20)
