# ─────────────────────────────────────────────
#  tools/histogram.py  —  Data Visualizer
# ─────────────────────────────────────────────

from collections import Counter
from tkinter import messagebox
import tkinter as tk

import state
from ui.mapping import build_shuttle_frame


def histogram_gui():
    if not state.excel_data or not state.excel_headers:
        messagebox.showerror("Error", "Load Excel data terlebih dahulu!")
        return

    sel_win = tk.Toplevel(state.root)
    sel_win.title("Histogram — Select Columns")
    sel_win.geometry("700x520")
    sel_win.transient(state.root)
    sel_win.grab_set()

    tk.Label(sel_win,
             text="Select the columns to include in this histogram.\n"
                  "All selected columns will be pooled into one frequency chart.",
             font=("Arial", 9, "italic"), fg="gray").pack(pady=(10, 4))

    list_available, list_selected = build_shuttle_frame(
        sel_win,
        label_left="Available Columns",
        label_right="Selected Columns",
        label_color="blue"
    )
    for col in state.excel_headers:
        list_available.insert(tk.END, col)

    order_frame = tk.LabelFrame(sel_win, text=" X-Axis Label Order (optional) ",
                                padx=8, pady=6)
    order_frame.pack(fill="x", padx=10, pady=6)

    tk.Label(order_frame,
             text="Comma-separated order of values, lowest → highest.\n"
                  "Leave blank to auto-sort (numeric: 1→N, text: alphabetical).",
             font=("Arial", 8), fg="gray").pack(anchor="w")

    order_var = tk.StringVar()
    tk.Entry(order_frame, textvariable=order_var, width=80).pack(fill="x", pady=4)

    tk.Label(order_frame,
             text='Example:  "Sangat Tidak Setuju, Tidak Setuju, Netral, Setuju, Sangat Setuju"'
                  '  or  "1, 2, 3, 4, 5"',
             font=("Arial", 7), fg="gray").pack(anchor="w")

    def confirm_selection():
        selected_cols = list(list_selected.get(0, tk.END))
        if not selected_cols:
            messagebox.showwarning("Warning", "Pilih minimal satu kolom!")
            return
        raw_order    = order_var.get().strip()
        custom_order = [x.strip() for x in raw_order.split(",") if x.strip()] if raw_order else None
        sel_win.destroy()
        _draw_histogram_window(selected_cols, custom_order)

    tk.Button(sel_win, text="Generate Histogram →",
              command=confirm_selection,
              bg="blue", fg="white", font=("Arial", 10, "bold"),
              height=2).pack(fill="x", padx=10, pady=10)

    sel_win.wait_window()


def _draw_histogram_window(selected_cols: list, custom_order):
    # ── Collect values ──────────────────────────────────────────────────
    all_values = []
    for col in selected_cols:
        for row in state.excel_data:
            v = row.get(col)
            if v is not None and str(v).strip() != "":
                all_values.append(v)

    if not all_values:
        messagebox.showinfo("Info", "Tidak ada data pada kolom yang dipilih.")
        return

    # ── Numeric or text ─────────────────────────────────────────────────
    coerced    = []
    is_numeric = True
    for v in all_values:
        try:
            coerced.append(int(float(str(v))))
        except (ValueError, TypeError):
            is_numeric = False
            break

    if is_numeric:
        all_values = coerced

    # ── Build categories ────────────────────────────────────────────────
    if custom_order:
        if is_numeric:
            try:
                categories = [int(x) for x in custom_order]
            except ValueError:
                categories = custom_order
        else:
            categories = custom_order
        present    = set(str(v) for v in all_values)
        covered    = set(str(c) for c in categories)
        extras     = sorted(present - covered)
        categories = categories + ([int(x) if is_numeric else x for x in extras])
    else:
        if is_numeric:
            mn, mx     = min(all_values), max(all_values)
            categories = list(range(mn, mx + 1))
        else:
            categories = sorted(set(str(v) for v in all_values))

    # ── Count ───────────────────────────────────────────────────────────
    counter  = Counter(all_values if is_numeric else [str(v) for v in all_values])
    cat_keys = categories if is_numeric else [str(c) for c in categories]
    counts   = [counter.get(k, 0) for k in cat_keys]
    total    = sum(counts)
    n_bars   = len(categories)

    # ── Chart window ────────────────────────────────────────────────────
    n_col_str = f"{len(selected_cols)} column{'s' if len(selected_cols) > 1 else ''}"
    chart_win = tk.Toplevel(state.root)
    chart_win.title(f"Histogram — {n_col_str}")
    chart_win.geometry("860x560")
    chart_win.resizable(True, True)

    # Title label
    col_list_str = ", ".join(selected_cols) if len(selected_cols) <= 4 \
                   else ", ".join(selected_cols[:4]) + f" … (+{len(selected_cols)-4} more)"
    tk.Label(chart_win,
             text=f"{n_col_str}: {col_list_str}",
             font=("Arial", 10, "bold")).pack(pady=(10, 0))
    tk.Label(chart_win,
             text=f"Total responses pooled: {total}",
             font=("Arial", 8), fg="gray").pack()

    # Canvas
    canvas = tk.Canvas(chart_win, bg="white")
    canvas.pack(fill="both", expand=True, padx=10, pady=10)

    def draw(event=None):
        canvas.delete("all")

        cw = canvas.winfo_width()
        ch = canvas.winfo_height()

        pad_left   = 55
        pad_right  = 20
        pad_top    = 20
        pad_bottom = 60

        chart_w = cw - pad_left - pad_right
        chart_h = ch - pad_top  - pad_bottom

        if chart_w <= 0 or chart_h <= 0 or n_bars == 0:
            return

        max_count = max(counts) if any(counts) else 1
        bar_w     = chart_w / n_bars
        gap       = bar_w * 0.2

        # Y-axis gridlines + labels
        y_steps = 5
        for i in range(y_steps + 1):
            val  = round(max_count * i / y_steps)
            y    = pad_top + chart_h - (chart_h * i / y_steps)
            canvas.create_line(pad_left, y, pad_left + chart_w, y,
                               fill="#e0e0e0", dash=(4, 4))
            canvas.create_text(pad_left - 6, y, text=str(val),
                               anchor="e", font=("Arial", 8), fill="#555")

        # Bars
        for i, (cat, count) in enumerate(zip(cat_keys, counts)):
            x1 = pad_left + i * bar_w + gap / 2
            x2 = pad_left + (i + 1) * bar_w - gap / 2
            bar_h = (count / max_count) * chart_h if max_count > 0 else 0
            y1 = pad_top + chart_h - bar_h
            y2 = pad_top + chart_h

            # Bar rectangle
            canvas.create_rectangle(x1, y1, x2, y2,
                                    fill="#4a90d9", outline="#2c5f8a", width=1)

            # Count on top
            canvas.create_text((x1 + x2) / 2, y1 - 8,
                               text=str(count),
                               font=("Arial", 8, "bold"), fill="#333")

            # Percentage inside bar (only if tall enough)
            if count > 0 and bar_h > 22 and total > 0:
                pct = count / total * 100
                canvas.create_text((x1 + x2) / 2, (y1 + y2) / 2,
                                   text=f"{pct:.1f}%",
                                   font=("Arial", 8, "bold"), fill="white")

            # X-axis label
            label = str(cat)
            canvas.create_text((x1 + x2) / 2, y2 + 14,
                               text=label, font=("Arial", 8),
                               fill="#333", angle=30 if len(label) > 6 else 0,
                               anchor="n" if len(label) > 6 else "center")

        # Axes
        canvas.create_line(pad_left, pad_top,
                           pad_left, pad_top + chart_h,
                           fill="#333", width=2)
        canvas.create_line(pad_left, pad_top + chart_h,
                           pad_left + chart_w, pad_top + chart_h,
                           fill="#333", width=2)

    canvas.bind("<Configure>", draw)