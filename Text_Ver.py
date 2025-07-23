import json
import tkinter as tk
from tkinter import ttk, messagebox
import re
from deepdiff import DeepDiff
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
import os

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
            else: # key is a string
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

def format_full_output(data):
    if not isinstance(data, dict):
        return json.dumps(data, indent=2, ensure_ascii=False)
    output_lines = []
    if "promoInfo" in data and isinstance(data["promoInfo"], list):
        # Sort promos robustly, handling non-digit promoNumbers
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
        # Use json.dumps for the key's value for consistent formatting
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
            label_result.config(text="✅ Text copied to clipboard!", foreground="#66ff99")
            root.after(1500, clear_label_result)
        except Exception as e:
            label_result.config(text=f"❌ Copy failed: {e}", foreground="#ff6666")
            root.after(1500, clear_label_result)
    else:
        label_result.config(text="⚠️ Nothing to copy.", foreground="#ffaa00")
        root.after(1500, clear_label_result)

def add_right_click_menu(widget):
    menu = tk.Menu(widget, tearoff=0, bg="#2e2e2e", fg="#f8f8f2")
    menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"))
    widget.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))

def bind_scroll(widget):
    widget.bind("<Enter>", lambda e: widget.bind_all("<MouseWheel>", lambda ev: widget.yview_scroll(int(-1*(ev.delta/120)), "units")))
    widget.bind("<Leave>", lambda e: widget.unbind_all("<MouseWheel>"))

def bind_paste_shortcuts(widget):
    def do_paste(event):
        widget.event_generate("<<Paste>>")
        return "break"
    for seq in ("<Control-v>", "<Control-V>", "<Shift-Insert>"):
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
    # This function uses a heuristic to highlight lines containing a changed key.
    # It may highlight extra lines if the key name is not unique within the object.
    text_widget.tag_remove("diff_highlight", "1.0", tk.END)
    text_widget.tag_configure("diff_highlight", foreground="#F700FF", font=("Segoe UI", 10, "bold"))
    for path in diff_paths:
        keys = re.findall(r"\['([^]]+)'\]|\[(\d+)\]", path)
        if not keys:
            continue
        last_key = keys[-1][0] or keys[-1][1]
        start = "1.0"
        while True:
            pos = text_widget.search(f'"{last_key}"', start, stopindex=tk.END)
            if not pos:
                break
            # Highlight the entire line where the key is found
            line_start = f"{pos.split('.')[0]}.0"
            line_end = f"{pos.split('.')[0]}.end"
            text_widget.tag_add("diff_highlight", line_start, line_end)
            start = line_end

# ----------------- Global Variables -----------------
EXPORT_FOLDER = os.path.join(os.getcwd(), "export")
EXCEL_PATH = os.path.join(EXPORT_FOLDER, "Compare_Export.xlsx")

# Create export folder if it doesn't exist
if not os.path.exists(EXPORT_FOLDER):
    try:
        os.makedirs(EXPORT_FOLDER)
    except Exception as e:
        messagebox.showerror("Folder Creation Failed", f"Could not create the export folder:\n{e}")
        raise

last_export_data = None # Store the latest comparison result

# ----------------- Excel Export Utility -----------------
def export_to_excel():
    if not last_export_data:
        messagebox.showwarning("No Comparison Data", "Please compare JSON files before exporting.")
        return

    base_part, comp_part = last_export_data
    diff_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Get single-line JSON strings from the input boxes for logging
    base_json_str = text_base.get("1.0", "end").strip().replace("\n", " ")
    compare_json_str = text_compare.get("1.0", "end").strip().replace("\n", " ")

    try:
        if os.path.exists(EXCEL_PATH):
            wb = load_workbook(EXCEL_PATH)
        else:
            wb = Workbook()
            wb.remove(wb.active)
    except Exception as e:
        messagebox.showerror("Excel Load Error", str(e))
        return

    # Delete old sheet and create a new one
    if "Comparison" in wb.sheetnames:
        del wb["Comparison"]
    ws = wb.create_sheet("Comparison")

    # --- Write input data to the first few rows ---
    ws.cell(row=1, column=1, value="Input: JSON Base (Onlinepro.json)")
    ws.cell(row=1, column=2, value="Input: JSON Compare (NewPro.json)")
    ws.cell(row=2, column=1, value=base_json_str)
    ws.cell(row=2, column=2, value=compare_json_str)

    # --- Write comparison result headers ---
    ws.cell(row=4, column=1, value="Result: JSON Base (Differences)")
    ws.cell(row=4, column=2, value="Result: JSON Compare (Differences)")

    current_row = 5

    def json_to_lines(data):
        if data is None: return []
        if isinstance(data, (dict, list)):
            return json.dumps(data, indent=2, ensure_ascii=False).splitlines()
        return [json.dumps(data, ensure_ascii=False)]

    base_promos = {p["promoNumber"]: p for p in base_part.get("promoInfo", []) if "promoNumber" in p}
    comp_promos = {p["promoNumber"]: p for p in comp_part.get("promoInfo", []) if "promoNumber" in p}
    all_promo_nums = sorted(
        set(base_promos) | set(comp_promos),
        key=lambda x: int(x) if str(x).isdigit() else float('inf')
    )

    for promo_num in all_promo_nums:
        ws.cell(row=current_row, column=1, value=f"========== promoNumber: {promo_num} ==========")
        current_row += 1

        base_obj = base_promos.get(promo_num, {})
        comp_obj = comp_promos.get(promo_num, {})

        base_lines = json_to_lines(base_obj)
        comp_lines = json_to_lines(comp_obj)
        max_lines = max(len(base_lines), len(comp_lines))

        for i in range(max_lines):
            b_line = base_lines[i] if i < len(base_lines) else ""
            c_line = comp_lines[i] if i < len(comp_lines) else ""
            
            cell_b = ws.cell(row=current_row, column=1, value=b_line)
            cell_c = ws.cell(row=current_row, column=2, value=c_line)

            if b_line != c_line:
                if b_line: cell_b.fill = diff_fill
                if c_line: cell_c.fill = diff_fill
            current_row += 1
        current_row += 1

    # Handle other fields outside of promoInfo
    others_base = {k: v for k, v in base_part.items() if k != "promoInfo"}
    others_comp = {k: v for k, v in comp_part.items() if k != "promoInfo"}

    if others_base or others_comp:
        ws.cell(row=current_row, column=1, value="========== Others ==========")
        current_row += 1

        all_keys = sorted(set(others_base.keys()) | set(others_comp.keys()))
        for key in all_keys:
            ws.cell(row=current_row, column=1, value=f'"{key}":')
            current_row += 1

            base_val = others_base.get(key)
            comp_val = others_comp.get(key)
            
            base_lines = json_to_lines(base_val)
            comp_lines = json_to_lines(comp_val)
            max_lines = max(len(base_lines), len(comp_lines))

            for i in range(max_lines):
                b_line = base_lines[i] if i < len(base_lines) else ""
                c_line = comp_lines[i] if i < len(comp_lines) else ""

                cell_b = ws.cell(row=current_row, column=1, value=b_line)
                cell_c = ws.cell(row=current_row, column=2, value=c_line)

                if b_line != c_line:
                    if b_line: cell_b.fill = diff_fill
                    if c_line: cell_c.fill = diff_fill
                current_row += 1
            current_row += 1

    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 70

    try:
        wb.save(EXCEL_PATH)
        messagebox.showinfo("Export Successful", f"Excel file saved to:\n{EXCEL_PATH}")
    except PermissionError:
        messagebox.showerror("Save Failed", "Permission denied. Please close the Excel file and try again.")
    except Exception as e:
        messagebox.showerror("Save Failed", f"An unexpected error occurred:\n{e}")
        
# ----------------- Core Function: compare_json -----------------
def compare_json():
    try:
        base_data = json.loads(text_base.get("1.0", tk.END))
        compare_data = json.loads(text_compare.get("1.0", tk.END))
    except json.JSONDecodeError as e:
        messagebox.showerror("Invalid JSON Format", str(e))
        return
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred while reading input: {e}")
        return

    remove_description(base_data)
    remove_description(compare_data)

    base_filtered = filter_out_debug(base_data)
    compare_filtered = filter_out_debug(compare_data)

    # --- Compare PromoInfo Section ---
    base_promos = {p["promoNumber"]: p for p in base_filtered.get("promoInfo", []) if "promoNumber" in p}
    compare_promos = {p["promoNumber"]: p for p in compare_filtered.get("promoInfo", []) if "promoNumber" in p}

    partial_base_result = {"promoInfo": []}
    partial_compare_result = {"promoInfo": []}
    total_diff_paths = []

    # Robustly sort promo numbers, handling non-integer values
    all_promo_numbers = sorted(
        set(base_promos.keys()) | set(compare_promos.keys()),
        key=lambda x: int(x) if str(x).isdigit() else float('inf')
    )

    for promo_num in all_promo_numbers:
        base_promo = base_promos.get(promo_num)
        compare_promo = compare_promos.get(promo_num)
        
        diff = DeepDiff(base_promo, compare_promo, ignore_order=False, report_repetition=True, view="tree")
        
        path_list = []
        if diff:
            for section in diff:
                for change in diff[section]:
                    if hasattr(change, 'path'):
                        path = change.path(output_format='list')
                        s = "".join(f"[{p}]" if isinstance(p, int) else f"['{p}']" for p in path)
                        path_list.append(s)
            
            total_diff_paths.extend([f"['promoInfo'][{len(partial_base_result['promoInfo'])}]{p}" for p in path_list])

        partial_base = build_partial_json(base_promo, path_list) if base_promo else {}
        partial_base["promoNumber"] = promo_num
        
        partial_compare = build_partial_json(compare_promo, path_list) if compare_promo else {}
        partial_compare["promoNumber"] = promo_num

        # Only add to results if there was a difference or if the promo is unique to one side
        if diff:
            partial_base_result["promoInfo"].append(partial_base)
            partial_compare_result["promoInfo"].append(partial_compare)

    # --- Compare Other Keys Section (All at once for correctness) ---
    others_base = {k: v for k, v in base_filtered.items() if k != 'promoInfo'}
    others_compare = {k: v for k, v in compare_filtered.items() if k != 'promoInfo'}
    
    other_diff = DeepDiff(others_base, others_compare, ignore_order=False, report_repetition=True, view="tree")
    
    if other_diff:
        path_list = []
        for section in other_diff:
            for change in other_diff[section]:
                if hasattr(change, 'path'):
                    path = change.path(output_format='list')
                    s = "".join(f"[{p}]" if isinstance(p, int) else f"['{p}']" for p in path)
                    # Reconstruct path from root key
                    if path:
                        path_list.append(f"['{path[0]}']" + s[s.find(']'):])
        
        total_diff_paths.extend(path_list)

        # Build partial JSONs for the 'other' keys and update the main result
        partial_base_others = build_partial_json(base_filtered, path_list)
        partial_compare_others = build_partial_json(compare_filtered, path_list)
        
        partial_base_result.update(partial_base_others)
        partial_compare_result.update(partial_compare_others)

    # --- Display Results ---
    base_result_str = format_full_output(partial_base_result)
    compare_result_str = format_full_output(partial_compare_result)

    text_partial_base.delete("1.0", tk.END)
    text_partial_compare.delete("1.0", tk.END) # <<< THIS LINE IS NOW FIXED
    text_partial_base.insert(tk.END, base_result_str)
    text_partial_compare.insert(tk.END, compare_result_str)

    highlight_promo_lines(text_partial_base)
    highlight_promo_lines(text_partial_compare)
    highlight_differences(text_partial_base, total_diff_paths)
    highlight_differences(text_partial_compare, total_diff_paths)

    global last_export_data
    last_export_data = (partial_base_result, partial_compare_result)

    label_result.config(text=f"🔍 Found {len(total_diff_paths)} differences.", foreground="#f8f8f2")

# ----------------- GUI Setup -----------------
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

root.grid_rowconfigure(5, weight=1) # Make row 5 (output) expand
root.grid_columnconfigure(0, weight=1)

top_frame = ttk.Frame(root)
top_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(10, 5))
top_frame.columnconfigure(0, weight=1)
ttk.Label(top_frame, text="🧠 JSON Compare Tool", style="Header.TLabel").pack()

frame_input = ttk.Frame(root)
frame_input.grid(row=1, column=0, sticky="nsew", padx=10)
frame_input.grid_columnconfigure(0, weight=1)
frame_input.grid_columnconfigure(1, weight=1)
frame_input.grid_rowconfigure(1, weight=1) # Allow text boxes to expand vertically

ttk.Label(frame_input, text="📘 JSON Base (Onlinepro.json)", style="Header.TLabel").grid(row=0, column=0, sticky="w")
text_base = tk.Text(frame_input, bg=TEXTBOX_BG, fg=DARK_TEXT, insertbackground="white", relief="groove", height=10)
text_base.grid(row=1, column=0, sticky="nsew", padx=(0, 5))
add_right_click_menu(text_base)
bind_scroll(text_base)
bind_paste_shortcuts(text_base)

ttk.Label(frame_input, text="📙 JSON Compare (NewPro.json)", style="Header.TLabel").grid(row=0, column=1, sticky="w")
text_compare = tk.Text(frame_input, bg=TEXTBOX_BG, fg=DARK_TEXT, insertbackground="white", relief="groove", height=10)
text_compare.grid(row=1, column=1, sticky="nsew", padx=(5, 0))
add_right_click_menu(text_compare)
bind_scroll(text_compare)
bind_paste_shortcuts(text_compare)

ttk.Button(root, text="🔍 Compare JSON", command=compare_json).grid(row=2, column=0, pady=10)

label_result = ttk.Label(root, text="", background=DARK_BG, font=("Segoe UI", 12, "bold"))
label_result.grid(row=3, column=0, pady=5)

frame_controls = ttk.Frame(root)
frame_controls.grid(row=4, column=0, pady=5)
ttk.Button(frame_controls, text="📋 Copy Base Diff", command=lambda: copy_text(text_partial_base)).pack(side="left", padx=15)
ttk.Button(frame_controls, text="📤 Export to Excel", command=export_to_excel).pack(side="left", padx=15)
ttk.Button(frame_controls, text="📋 Copy Compare Diff", command=lambda: copy_text(text_partial_compare)).pack(side="left", padx=15)

frame_output = ttk.Frame(root)
frame_output.grid(row=5, column=0, sticky="nsew", padx=10, pady=(0, 10))
frame_output.grid_columnconfigure(0, weight=1)
frame_output.grid_columnconfigure(1, weight=1)
frame_output.grid_rowconfigure(1, weight=1)

ttk.Label(frame_output, text="📘 Base Differences", style="Header.TLabel").grid(row=0, column=0, sticky="w")
text_partial_base = tk.Text(frame_output, bg=TEXTBOX_BG, fg=DARK_TEXT, insertbackground="white", relief="ridge")
text_partial_base.grid(row=1, column=0, sticky="nsew", padx=(0, 5))
add_right_click_menu(text_partial_base)
bind_scroll(text_partial_base)
bind_paste_shortcuts(text_partial_base)

ttk.Label(frame_output, text="📙 Compare Differences", style="Header.TLabel").grid(row=0, column=1, sticky="w")
text_partial_compare = tk.Text(frame_output, bg=TEXTBOX_BG, fg=DARK_TEXT, insertbackground="white", relief="ridge")
text_partial_compare.grid(row=1, column=1, sticky="nsew", padx=(5, 0))
add_right_click_menu(text_partial_compare)
bind_scroll(text_partial_compare)
bind_paste_shortcuts(text_partial_compare)

root.mainloop()