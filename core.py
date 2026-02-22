# ─────────────────────────────────────────────
#  core.py  —  Random generation, URL compute,
#              bulk submission, undo/redo
# ─────────────────────────────────────────────

import math
import random
import copy
import time
import threading
import requests
from tkinter import messagebox
from urllib.parse import urlunparse, urlparse

import state


# ─────────────────────────────────────────────
#  Helper Function
# ─────────────────────────────────────────────

def _log(msg: str):
    state.preview_text.config(state="normal")
    state.preview_text.insert("end", msg)
    state.preview_text.see("end")
    state.preview_text.config(state="disabled")


# ─────────────────────────────────────────────
#  Random-Value Generator
# ─────────────────────────────────────────────

def _random_value(mini: int, maxi: int, method: str) -> int:
    if method == "True Random":
        return random.randint(mini, maxi)
    elif method == "Normal (Bell Curve)":
        raw = random.betavariate(5, 5)
    elif method == "Skewed Positive":
        raw = random.betavariate(2, 5)
    else:  # Skewed Negative
        raw = random.betavariate(5, 2)
    return round(mini + raw * (maxi - mini))


def _generate_population(mini: int, maxi: int, method: str, n: int) -> list:
    scale = maxi - mini + 1

    if method == "True Random":
        return [random.randint(mini, maxi) for _ in range(n)]

    if method == "Normal (Bell Curve)":
        mu = (scale - 1) / 2
        sigma = scale / 4
        weights = [math.exp(-0.5 * ((i - mu) / sigma) ** 2) for i in range(scale)]
    elif method == "Skewed Positive":
        weights = [((i + 1) / scale) ** 0.5 for i in range(scale)]
    else:  # Skewed Negative
        weights = [((scale - i) / scale) ** 0.5 for i in range(scale)]

    total_w = sum(weights)
    cumulative = []
    running = 0
    for w in weights:
        running += w / total_w
        cumulative.append(running)

    def weighted_sample():
        r = random.random()
        for i, threshold in enumerate(cumulative):
            if r <= threshold:
                return mini + i
        return maxi

    return [weighted_sample() for _ in range(n)]


# ─────────────────────────────────────────────
#  URL Generation
# ─────────────────────────────────────────────

def _compute_urls() -> list:
    mapping = {eid: var.get().strip() for eid, var in state.mapping_vars.items()}
    urls = []
    for row in state.excel_data:
        current_query = {k: v[:] for k, v in state.query_dict.items()}
        for entry_id, col in mapping.items():
            val = "" if col == "NULL" else row.get(col, "")
            current_query[entry_id] = [str(val) if val is not None else ""]

        page_history_value = current_query.get("pageHistory", ["0"])[0]
        query_str = f"pageHistory={page_history_value}"
        for key, values in current_query.items():
            if key in ("pageHistory", "usp"):
                continue
            for v in values:
                query_str += f"&{key}={v}"

        urls.append(urlunparse((
            state.parsed_url.scheme, state.parsed_url.netloc, state.parsed_url.path,
            state.parsed_url.params, query_str, state.parsed_url.fragment
        )))
    return urls


def _update_url_ui(urls: list):
    state.final_output_urls = urls

    state.preview_text.config(state="normal")
    state.preview_text.delete("1.0", "end")
    state.preview_text.insert("end", f"> Total URLs Generated: {len(urls)}\n> Ready for submission.\n\n")
    for u in urls[:10]:
        p = urlparse(u)
        state.preview_text.insert("end", f"{p.path}?{p.query}\n\n")
    if len(urls) > 10:
        state.preview_text.insert("end", f"... and {len(urls) - 10} more rows")
    state.preview_text.config(state="disabled")
    state.progress_var.set(100)
    state.root.after(1000, lambda: state.progress_var.set(0))


def generate_urls():
    if not state.mapping_vars:
        messagebox.showerror("Error", "Mapping belum dibuat!")
        return
    _update_url_ui(_compute_urls())


def generate_urls_threaded():
    if not state.mapping_vars:
        messagebox.showerror("Error", "Mapping belum dibuat!")
        return

    state.preview_text.config(state="normal")
    state.preview_text.delete("1.0", "end")
    state.preview_text.insert("end", "> Generating URLs...\n")
    state.preview_text.config(state="disabled")

    def worker():
        try:
            urls = _compute_urls()
            state.root.after(0, lambda: _update_url_ui(urls))
        except Exception as e:
            state.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            state.root.after(0, lambda: _log(f"> Error generating URLs: {e}\n"))

    threading.Thread(target=worker, daemon=True).start()


# ─────────────────────────────────────────────
#  Bulk Submission
# ─────────────────────────────────────────────

def start_bulk_submission():
    if state.is_submitting:
        return
    threading.Thread(target=execute_submissions, daemon=True).start()


def execute_submissions():
    if not state.final_output_urls:
        messagebox.showerror("Error", "Generate URL dulu!")
        return
    if not messagebox.askyesno("Konfirmasi", f"Kirim {len(state.final_output_urls)} URL?"):
        return

    state.is_submitting = True
    success = fail = 0

    # Clear log and write header
    state.preview_text.config(state="normal")
    state.preview_text.delete("1.0", "end")
    state.preview_text.insert("end", f"> Submitting {len(state.final_output_urls)} rows...\n\n")
    state.preview_text.config(state="disabled")

    with requests.Session() as session:
        for idx, url in enumerate(state.final_output_urls):
            if not state.is_submitting:
                state.preview_text.config(state="normal")
                state.preview_text.insert("end", f"> Stopped at row {idx + 1}.\n")
                state.preview_text.config(state="disabled")
                break
            try:
                response = session.get(url, timeout=10)
                if response.status_code == 200:
                    success += 1
                    msg = f"Row {idx + 1}: OK\n"
                else:
                    fail += 1
                    msg = f"Row {idx + 1}: FAILED — HTTP {response.status_code}\n"
            except Exception as e:
                fail += 1
                msg = f"Row {idx + 1}: ERROR — {e}\n"

            state.preview_text.config(state="normal")
            state.preview_text.insert("end", msg)
            state.preview_text.see("end")
            state.preview_text.config(state="disabled")
            state.progress_var.set((idx + 1) / len(state.final_output_urls) * 100)
            state.root.update_idletasks()
            time.sleep(0.7)

    state.is_submitting = False
    summary = f"\n> Done! Berhasil: {success}  Gagal: {fail}\n"
    state.preview_text.config(state="normal")
    state.preview_text.insert("end", summary)
    state.preview_text.see("end")
    state.preview_text.config(state="disabled")
    state.progress_var.set(0)
    messagebox.showinfo("Done", f"Selesai!\nBerhasil: {success}  Gagal: {fail}")


def stop_submission():
    state.is_submitting = False
    state.preview_text.config(state="normal")
    state.preview_text.insert("end", "> Submission stop requested by user.\n")
    state.preview_text.see("end")
    state.preview_text.config(state="disabled")
    messagebox.showwarning("Stop", "Submission akan dihentikan setelah baris ini selesai.")


# ─────────────────────────────────────────────
#  Undo / Redo
# ─────────────────────────────────────────────

def save_undo_state():
    s = {
        "excel_data": copy.deepcopy(state.excel_data),
        "mapping":    {eid: var.get() for eid, var in state.mapping_vars.items()}
    }
    state.undo_stack.append(s)
    if len(state.undo_stack) > state.MAX_UNDO:
        state.undo_stack.pop(0)
    state.redo_stack.clear()


def undo(event=None):
    if not state.undo_stack:
        return
    current = {
        "excel_data": copy.deepcopy(state.excel_data),
        "mapping":    {eid: var.get() for eid, var in state.mapping_vars.items()}
    }
    state.redo_stack.append(current)
    last = state.undo_stack.pop()
    state.excel_data = last["excel_data"]
    for eid, val in last["mapping"].items():
        if eid in state.mapping_vars:
            state.mapping_vars[eid].set(val)
    from ui.spreadsheet import update_spreadsheet_view
    update_spreadsheet_view()


def redo(event=None):
    if not state.redo_stack:
        return
    current = {
        "excel_data": copy.deepcopy(state.excel_data),
        "mapping":    {eid: var.get() for eid, var in state.mapping_vars.items()}
    }
    state.undo_stack.append(current)
    nxt = state.redo_stack.pop()
    state.excel_data = nxt["excel_data"]
    for eid, val in nxt["mapping"].items():
        if eid in state.mapping_vars:
            state.mapping_vars[eid].set(val)
    from ui.spreadsheet import update_spreadsheet_view
    update_spreadsheet_view()
