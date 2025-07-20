import json
import tkinter as tk
from tkinter import ttk, messagebox
import re
import os
from deepdiff import DeepDiff
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment

# ----------------- JSON Utility -----------------
def remove_description(obj):
    if isinstance(obj, dict):
        obj.pop("description", None)
        for v in obj.values():
            remove_description(v)
    elif isinstance(obj, list):
        for item in obj:
            remove_description(item)

def filter_out_debug(obj):
    if isinstance(obj, dict):
        return {k: filter_out_debug(v) for k, v in obj.items() if "debug" not in k.lower()}
    elif isinstance(obj, list):
        return [filter_out_debug(i) for i in obj]
    else:
        return obj

def build_partial_json(base, diff_paths):
    partial = {}
    for path in diff_paths:
        keys = re.findall(r"\['([^]]+)'\]|\[(\d+)\]", path)
        keys = [k[0] if k[0] else int(k[1]) for k in keys]
        current_src = base
        current_partial = partial
        parents = []
        for i, key in enumerate(keys):
            is_last = (i == len(keys) - 1)
            if isinstance(current_src, dict) and key not in current_src:
                break
            if isinstance(current_src, list) and (not isinstance(key, int) or key >= len(current_src)):
                break
            if isinstance(key, int):
                if not isinstance(current_partial, list):
                    if isinstance(current_partial, dict) and not current_partial:
                        new_list = []
                        if parents:
                            parent, parent_key = parents[-1]
                            parent[parent_key] = new_list
                        else:
                            partial = new_list
                        current_partial = new_list
                    else:
                        break
                while len(current_partial) <= key:
                    current_partial.append({})
                if is_last:
                    current_partial[key] = current_src[key]
                else:
                    parents.append((current_partial, key))
                    current_partial = current_partial[key]
                    current_src = current_src[key]
            else:
                if not isinstance(current_partial, dict):
                    break
                if key not in current_partial:
                    current_partial[key] = {}
                if is_last:
                    current_partial[key] = current_src[key]
                else:
                    parents.append((current_partial, key))
                    current_partial = current_partial[key]
                    current_src = current_src[key]
    return partial

def fill_missing_promo_numbers(partial, full):
    if not ("promoInfo" in partial and isinstance(partial["promoInfo"], list)):
        return
    if not ("promoInfo" in full and isinstance(full["promoInfo"], list)):
        return
    full_promos = full["promoInfo"]
    partial_promos = partial["promoInfo"]
    for i, promo_partial in enumerate(partial_promos):
        if "promoNumber" not in promo_partial:
            promo_partial["promoNumber"] = full_promos[i].get("promoNumber", "N/A") if i < len(full_promos) else "N/A"

def format_full_output(data):
    if not isinstance(data, dict):
        return json.dumps(data, indent=2, ensure_ascii=False)
    output_lines = []
    if "promoInfo" in data and isinstance(data["promoInfo"], list):
        sorted_promos = sorted(
            data["promoInfo"],
            key=lambda p: int(p.get("promoNumber", "0")) if str(p.get("promoNumber", "0")).isdigit() else float('inf')
        )
        for promo in sorted_promos:
            promo_number = promo.get("promoNumber", "N/A")
            output_lines.append(f"========== promoNumber: {promo_number} ==========")
            output_lines.append(json.dumps(promo, indent=2, ensure_ascii=False))
            output_lines.append("")
    for key, value in data.items():
        if key == "promoInfo":
            continue
        output_lines.append(f'"{key}": {json.dumps(value, indent=2, ensure_ascii=False)}')
        output_lines.append("")
    return "\n".join(output_lines).strip()

# ----------------- GUI Utility -----------------
def clear_label_result():
    label_result.config(text="")

def copy_text(widget):
    content = widget.get("1.0", tk.END).strip()
    if content:
        try:
            root.clipboard_clear()
            root.clipboard_append(content)
            label_result.config(text="✅ ข้อความถูกคัดลอกไปยังคลิปบอร์ดแล้ว", foreground="#66ff99")
            label_result.after(1000, clear_label_result)
        except Exception as e:
            label_result.config(text=f"❌ ไม่สามารถคัดลอกข้อความได้: {e}", foreground="#ff6666")
            label_result.after(1000, clear_label_result)
    else:
        label_result.config(text="⚠️ ไม่มีข้อความให้คัดลอก", foreground="#ffaa00")
        label_result.after(1000, clear_label_result)

def add_right_click_menu(widget):
    menu = tk.Menu(widget, tearoff=0, bg="#2e2e2e", fg="#f8f8f2")
    menu.add_command(label="วาง (Paste)", command=lambda: widget.event_generate("<<Paste>>"))
    widget.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))

def bind_scroll(widget):
    widget.bind("<Enter>", lambda e: widget.bind_all("<MouseWheel>", lambda ev: widget.yview_scroll(int(-1*(ev.delta/120)), "units")))
    widget.bind("<Leave>", lambda e: widget.unbind_all("<MouseWheel>"))

def bind_paste_shortcuts(widget):
    def do_paste(event):
        widget.event_generate("<<Paste>>")
        return "break"
    for seq in ("<Control-v>", "<Control-V>", "<Shift-Insert>", "<Control-Insert>"):
        widget.bind(seq, do_paste)

def highlight_promo_lines(text_widget):
    text_widget.tag_configure("highlight", foreground="#00ff00", font=("Segoe UI", 10, "bold"))
    start = "1.0"
    while True:
        pos = text_widget.search(r"^=+ promoNumber: .* =+$", start, stopindex=tk.END, regexp=True)
        if not pos:
            break
        text_widget.tag_add("highlight", pos, f"{pos} lineend")
        start = f"{pos} lineend"

def highlight_differences(text_widget, diff_paths):
    text_widget.tag_remove("diff_highlight", "1.0", tk.END)
    text_widget.tag_configure("diff_highlight", foreground="#F700FF", font=("Segoe UI", 10, "bold"))
    for path in diff_paths:
        keys = re.findall(r"\['([^]]+)'\]|\[(\d+)\]", path)
        last_key = keys[-1][0] if keys and keys[-1][0] else (keys[-1][1] if keys else None)
        if not last_key:
            continue
        start = "1.0"
        while True:
            pos = text_widget.search(f'"{last_key}"', start, stopindex=tk.END)
            if not pos:
                break
            text_widget.tag_add("diff_highlight", f"{pos.split('.')[0]}.0", f"{pos.split('.')[0]}.end")
            start = f"{pos.split('.')[0]}.end"

# ----------------- Core Functions -----------------
def compare_json():
    try:
        base_data = json.loads(text_base.get("1.0", tk.END))
        compare_data = json.loads(text_compare.get("1.0", tk.END))
    except json.JSONDecodeError as e:
        messagebox.showerror("รูปแบบ JSON ไม่ถูกต้อง", str(e))
        return

    remove_description(base_data)
    remove_description(compare_data)
    base_filtered = filter_out_debug(base_data)
    compare_filtered = filter_out_debug(compare_data)
    diff = DeepDiff(base_filtered, compare_filtered, ignore_order=False, report_repetition=True, view="tree")
    if not diff:
        label_result.config(text="✅ ไม่มีความแตกต่างระหว่าง JSON ทั้งสองไฟล์")
        text_partial_base.delete("1.0", tk.END)
        text_partial_compare.delete("1.0", tk.END)
        return

    path_list = []
    for section in diff:
        for change in diff[section]:
            if hasattr(change, 'path'):
                p = change.path(output_format='list')
                path_list.append("".join(f"[{x}]" if isinstance(x, int) else f"['{x}']" for x in p))

    partial_base = build_partial_json(base_filtered, path_list)
    partial_compare = build_partial_json(compare_filtered, path_list)
    fill_missing_promo_numbers(partial_base, base_filtered)
    fill_missing_promo_numbers(partial_compare, compare_filtered)

    text_partial_base.delete("1.0", tk.END)
    text_partial_compare.delete("1.0", tk.END)
    text_partial_base.insert(tk.END, format_full_output(partial_base))
    text_partial_compare.insert(tk.END, format_full_output(partial_compare))

    highlight_promo_lines(text_partial_base)
    highlight_promo_lines(text_partial_compare)
    highlight_differences(text_partial_base, path_list)
    highlight_differences(text_partial_compare, path_list)

    total_diff = sum(len(diff[sec]) for sec in diff)
    label_result.config(text=f"🔍 พบความแตกต่างทั้งหมด {total_diff} จุด")

def export_to_excel():
    try:
        base_data = json.loads(text_base.get("1.0", tk.END))
        compare_data = json.loads(text_compare.get("1.0", tk.END))
    except json.JSONDecodeError as e:
        messagebox.showerror("รูปแบบ JSON ไม่ถูกต้อง", str(e))
        return

    remove_description(base_data)
    remove_description(compare_data)
    base_filtered = filter_out_debug(base_data)
    compare_filtered = filter_out_debug(compare_data)
    diff = DeepDiff(base_filtered, compare_filtered, ignore_order=False, report_repetition=True, view="tree")

    path_list = []
    for section in diff:
        for change in diff[section]:
            if hasattr(change, 'path'):
                p = change.path(output_format='list')
                path_list.append("".join(f"[{x}]" if isinstance(x, int) else f"['{x}']" for x in p))

    partial_base = build_partial_json(base_filtered, path_list)
    partial_compare = build_partial_json(compare_filtered, path_list)
    fill_missing_promo_numbers(partial_base, base_filtered)
    fill_missing_promo_numbers(partial_compare, compare_filtered)

    filename = os.path.expanduser(r"C:\Users\natth\OneDrive\compare_test.xlsx")
    sheet_name = "CompareDiff"

    if os.path.exists(filename):
        wb = load_workbook(filename)
    else:
        wb = Workbook()
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) == 1:
        wb.remove(wb["Sheet"])
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet(title=sheet_name)

    # Style
    # Highlight style: only value cell (not full row)
    highlight_fill_value = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")  # light orange
    header_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    promo_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    promos_base = partial_base.get("promoInfo", [])
    promos_compare = partial_compare.get("promoInfo", [])
    promo_dict = {}
    for promo in promos_base:
        promo_dict[promo.get("promoNumber", "N/A")] = {"base": promo, "compare": {}}
    for promo in promos_compare:
        promo_dict.setdefault(promo.get("promoNumber", "N/A"), {}).update({"compare": promo})

    def recursive_diff(base, compare, prefix=""):
        """Return list of (field, base_val, compare_val) for all fields (only those that differ)"""
        result = []
        if isinstance(base, dict) or isinstance(compare, dict):
            keys = set(base.keys() if isinstance(base, dict) else []).union(compare.keys() if isinstance(compare, dict) else [])
            for key in sorted(keys):
                b = base.get(key, "") if isinstance(base, dict) else ""
                c = compare.get(key, "") if isinstance(compare, dict) else ""
                result.extend(recursive_diff(b, c, f"{prefix}{key}."))
        elif isinstance(base, list) or isinstance(compare, list):
            max_len = max(len(base) if isinstance(base, list) else 0, len(compare) if isinstance(compare, list) else 0)
            for i in range(max_len):
                b = base[i] if isinstance(base, list) and i < len(base) else ""
                c = compare[i] if isinstance(compare, list) and i < len(compare) else ""
                result.extend(recursive_diff(b, c, f"{prefix}[{i}]."))
        else:
            if base != compare:
                result.append((prefix.rstrip("."), base, compare))
        return result

    row_idx = 1
    # Header row
    ws.append(["", "Key", "Base", "Compare"])
    for col in range(1, 5):
        ws.cell(row=row_idx, column=col).fill = header_fill
        ws.cell(row=row_idx, column=col).border = thin_border
        ws.cell(row=row_idx, column=col).alignment = Alignment(vertical='top')
    row_idx += 1

    for promo_num, data in promo_dict.items():
        # Header row: promoNumber (only col 1 has promoNumber, col 2/3 has promoNumber string)
        ws.append([promo_num, f"========== promoNumber: {promo_num} ==========", "", ""])
        for col in range(1, 5):
            ws.cell(row=row_idx, column=col).fill = promo_fill
            ws.cell(row=row_idx, column=col).border = thin_border
            ws.cell(row=row_idx, column=col).alignment = Alignment(vertical='top')
        row_idx += 1

        # Dump JSON (only diff part) for each side, but split key and value
        def parse_json_lines(json_lines):
            # Returns list of (indent, key, value) or (indent, '', line) for non-key lines
            result = []
            for line in json_lines:
                indent = len(line) - len(line.lstrip(' '))
                striped = line.strip()
                if striped.startswith('"') and ':' in striped:
                    key_part, val_part = striped.split(':', 1)
                    key = key_part.strip().strip('"')
                    value = val_part.strip().rstrip(',')
                    result.append((indent, key, value))
                else:
                    result.append((indent, '', striped))
            return result

        base_json = data.get("base", {})
        compare_json = data.get("compare", {})
        base_lines = json.dumps(base_json, indent=2, ensure_ascii=False).splitlines()
        compare_lines = json.dumps(compare_json, indent=2, ensure_ascii=False).splitlines()
        base_parsed = parse_json_lines(base_lines)
        compare_parsed = parse_json_lines(compare_lines)
        max_lines = max(len(base_parsed), len(compare_parsed))

        for i in range(max_lines):
            base_item = base_parsed[i] if i < len(base_parsed) else (0, '', '')
            compare_item = compare_parsed[i] if i < len(compare_parsed) else (0, '', '')
            indent = max(base_item[0], compare_item[0])
            key = base_item[1] or compare_item[1]
            base_val = base_item[2] if base_item[1] else ''
            compare_val = compare_item[2] if compare_item[1] else ''
            # If not a key-value line, show as structure
            if not key:
                ws.append(["", "", base_item[2], compare_item[2]])
            else:
                ws.append(["", ' ' * indent + key, base_val, compare_val])
            # Always set border and alignment
            for col in range(1, 5):
                ws.cell(row=row_idx, column=col).border = thin_border
                ws.cell(row=row_idx, column=col).alignment = Alignment(vertical='top')
            # Highlight only value cell if different
            if key and base_val != compare_val:
                if base_val != '':
                    ws.cell(row=row_idx, column=3).fill = highlight_fill_value
                if compare_val != '':
                    ws.cell(row=row_idx, column=4).fill = highlight_fill_value
            row_idx += 1
        # Blank row between promos
        ws.append(["", "", "", ""])
        row_idx += 1

    try:
        wb.save(filename)
        label_result.config(text=f"✅ บันทึกไฟล์ Excel แล้ว: {os.path.abspath(filename)}", foreground="#66ff99")
    except Exception as e:
        messagebox.showerror("Save Error", f"ไม่สามารถบันทึกไฟล์ได้: {e}")

# ----------------- GUI -----------------
root = tk.Tk()
root.title("🧠 JSON Compare Tool")
root.attributes("-fullscreen", True)

is_fullscreen = True
def toggle_fullscreen(event=None):
    global is_fullscreen
    is_fullscreen = not is_fullscreen
    root.attributes("-fullscreen", is_fullscreen)
def exit_fullscreen(event=None):
    root.attributes("-fullscreen", False)
root.bind("<F11>", toggle_fullscreen)
root.bind("<Escape>", exit_fullscreen)

DARK_BG = "#2e2e2e"
DARK_TEXT = "#f8f8f2"
TEXTBOX_BG = "#1e1e1e"
HIGHLIGHT = "#3c3f41"

root.configure(bg=DARK_BG)
style = ttk.Style()
style.theme_use("clam")
style.configure("TFrame", background=DARK_BG)
style.configure("TLabel", background=DARK_BG, foreground=DARK_TEXT)
style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"), background=DARK_BG, foreground=DARK_TEXT)
style.configure("TButton", background=HIGHLIGHT, foreground="#ffffff", relief="flat", padding=6)
style.map("TButton", background=[("active", "#505354")], foreground=[("active", "#ffffff")])
style.configure("TLabelframe", background=DARK_BG, foreground=DARK_TEXT)
style.configure("TLabelframe.Label", background=DARK_BG, foreground=DARK_TEXT)

root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

ttk.Label(root, text="🧠 JSON Compare Tool", style="Header.TLabel").grid(row=0, column=0, pady=(10,5))

frame_input = ttk.Frame(root)
frame_input.grid(row=1, column=0, sticky="nsew", padx=10)
frame_input.grid_columnconfigure(0, weight=1)
frame_input.grid_columnconfigure(1, weight=1)
frame_input.grid_rowconfigure(0, weight=1)

frame_compare = ttk.Frame(frame_input)
frame_compare.grid(row=0, column=0, padx=(0,5), sticky="nsew")
ttk.Label(frame_compare, text="📙 JSON Compare (NewPro.json)", style="Header.TLabel").pack(anchor="w")
text_compare = tk.Text(frame_compare, bg=TEXTBOX_BG, fg=DARK_TEXT, insertbackground="white", relief="groove")
text_compare.pack(fill="both", expand=True)
add_right_click_menu(text_compare)
bind_scroll(text_compare)
bind_paste_shortcuts(text_compare)

frame_base = ttk.Frame(frame_input)
frame_base.grid(row=0, column=1, padx=(5,0), sticky="nsew")
ttk.Label(frame_base, text="📘 JSON Base (Onlinepro.json)", style="Header.TLabel").pack(anchor="w")
text_base = tk.Text(frame_base, bg=TEXTBOX_BG, fg=DARK_TEXT, insertbackground="white", relief="groove")
text_base.pack(fill="both", expand=True)
add_right_click_menu(text_base)
bind_scroll(text_base)
bind_paste_shortcuts(text_base)

ttk.Button(root, text="🔍 เปรียบเทียบ JSON", command=compare_json).grid(row=2, column=0, pady=10)

label_result = ttk.Label(root, text="", background=DARK_BG, font=("Segoe UI", 12, "bold"))
label_result.grid(row=3, column=0, pady=5)

frame_copy = ttk.Frame(root)
frame_copy.grid(row=4, column=0)
ttk.Button(frame_copy, text="📋 Copy Compare Diff", command=lambda: copy_text(text_partial_compare)).pack(side="left", padx=15)
ttk.Button(frame_copy, text="📤 Export Excel", command=export_to_excel).pack(side="left", padx=15)
ttk.Button(frame_copy, text="📋 Copy Base Diff", command=lambda: copy_text(text_partial_base)).pack(side="left", padx=15)

frame_output = ttk.Frame(root)
frame_output.grid(row=5, column=0, sticky="nsew", padx=10, pady=(0,10))
frame_output.grid_columnconfigure(0, weight=1)
frame_output.grid_columnconfigure(1, weight=1)
frame_output.grid_rowconfigure(0, weight=1)

frame_diff_compare = ttk.Frame(frame_output)
frame_diff_compare.grid(row=0, column=0, sticky="nsew", padx=(0,5))
ttk.Label(frame_diff_compare, text="📙 JSON Compare - Differences", style="Header.TLabel").pack(anchor="w")
text_partial_compare = tk.Text(frame_diff_compare, bg=TEXTBOX_BG, fg=DARK_TEXT, insertbackground="white", relief="ridge")
text_partial_compare.pack(fill="both", expand=True)
add_right_click_menu(text_partial_compare)
bind_scroll(text_partial_compare)
bind_paste_shortcuts(text_partial_compare)

frame_diff_base = ttk.Frame(frame_output)
frame_diff_base.grid(row=0, column=1, sticky="nsew", padx=(5,0))
ttk.Label(frame_diff_base, text="📘 JSON Base - Differences", style="Header.TLabel").pack(anchor="w")
text_partial_base = tk.Text(frame_diff_base, bg=TEXTBOX_BG, fg=DARK_TEXT, insertbackground="white", relief="ridge")
text_partial_base.pack(fill="both", expand=True)
add_right_click_menu(text_partial_base)
bind_scroll(text_partial_base)
bind_paste_shortcuts(text_partial_base)

root.mainloop()
