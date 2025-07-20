import json
import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip
import re
from deepdiff import DeepDiff
from openpyxl import Workbook
from openpyxl.styles import PatternFill

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
    """
    เติม promoNumber ให้กับ promoInfo ที่ไม่มีใน partial จาก full (ต้นฉบับ)
    """
    if not ("promoInfo" in partial and isinstance(partial["promoInfo"], list)):
        return
    if not ("promoInfo" in full and isinstance(full["promoInfo"], list)):
        return

    full_promos = full["promoInfo"]
    partial_promos = partial["promoInfo"]

    for i, promo_partial in enumerate(partial_promos):
        if "promoNumber" not in promo_partial:
            if i < len(full_promos):
                promo_partial["promoNumber"] = full_promos[i].get("promoNumber", "N/A")
            else:
                promo_partial["promoNumber"] = "N/A"

def format_full_output(data):
    if not isinstance(data, dict):
        return json.dumps(data, indent=2, ensure_ascii=False)

    output_lines = []

    if "promoInfo" in data and isinstance(data["promoInfo"], list):
        # 🔽 Sort by promoNumber
        sorted_promos = sorted(
            data["promoInfo"],
            key=lambda p: int(p.get("promoNumber", "0")) if str(p.get("promoNumber", "0")).isdigit() else float('inf')
        )
        for promo in sorted_promos:
            promo_number = promo.get("promoNumber", "N/A")
            output_lines.append(f"========== promoNumber: {promo_number} ==========")
            # แสดงข้อมูลเต็มของแต่ละโปรโมชั่น
            promo_str = json.dumps(promo, indent=2, ensure_ascii=False)
            output_lines.append(promo_str)
            output_lines.append("")  # เว้นบรรทัด

    # สำหรับ key อื่นที่ไม่ใช่ promoInfo
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
    def popup(event):
        menu.tk_popup(event.x_root, event.y_root)
    widget.bind("<Button-3>", popup)

def bind_scroll(widget):
    widget.bind("<Enter>", lambda e: widget.bind_all("<MouseWheel>", lambda ev: widget.yview_scroll(int(-1*(ev.delta/120)), "units")))
    widget.bind("<Leave>", lambda e: widget.unbind_all("<MouseWheel>"))

def bind_paste_shortcuts(widget):
    def do_paste(event):
        widget.event_generate("<<Paste>>")
        return "break"
    widget.bind("<Control-v>", do_paste)
    widget.bind("<Control-V>", do_paste)
    widget.bind("<Shift-Insert>", do_paste)
    widget.bind("<Control-Insert>", do_paste)

def highlight_promo_lines(text_widget):
    text_widget.tag_configure("highlight", foreground="#00ff00", font=("Segoe UI", 10, "bold"))
    start = "1.0"
    while True:
        start = text_widget.search(r"^=+ promoNumber: .* =+$", start, stopindex=tk.END, regexp=True)
        if not start:
            break
        end = f"{start} lineend"
        text_widget.tag_add("highlight", start, end)
        start = end

def highlight_differences(text_widget, diff_paths):
    # เคลียร์ tag ก่อน
    text_widget.tag_remove("diff_highlight", "1.0", tk.END)

    # ตั้งค่า tag
    # text_widget.tag_configure("diff_highlight", background="#FFFFFF")  # สีไฮไลท์
    text_widget.tag_configure("diff_highlight", foreground="#F700FF", font=("Segoe UI", 10, "bold"))  # สีตัวหนังสือไฮไลท์
    text_widget.tag_configure("diff_highlight", font=("Segoe UI", 10, "bold"))  # ตัวหนา

    for path in diff_paths:
        # แปลง path deepdiff -> regex pattern สำหรับค้นหาใน Text widget
        # ตัวอย่าง path: "['promoInfo'][0]['accounts'][1]['endingBalance']"
        # เราจะหา key-value คู่ที่อยู่ใกล้ path นี้ใน text json
        
        # ดึง key สุดท้ายใน path เพื่อใช้ค้นหา
        keys = re.findall(r"\['([^]]+)'\]|\[(\d+)\]", path)
        last_key = None
        if keys:
            last_key = keys[-1][0] if keys[-1][0] else keys[-1][1]

        if not last_key:
            continue
        
        # ค้นหาบรรทัดที่มี key หรือ key พร้อมค่าที่น่าสนใจ
        start_index = "1.0"
        while True:
            pos = text_widget.search(f'"{last_key}"', start_index, stopindex=tk.END)
            if not pos:
                break
            # ไฮไลท์ทั้งบรรทัด
            line_start = f"{pos.split('.')[0]}.0"
            line_end = f"{pos.split('.')[0]}.end"
            text_widget.tag_add("diff_highlight", line_start, line_end)
            start_index = line_end

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
                path = change.path(output_format='list')
                s = "".join(f"[{p}]" if isinstance(p, int) else f"['{p}']" for p in path)
                path_list.append(s)

    partial_base = build_partial_json(base_filtered, path_list)
    partial_compare = build_partial_json(compare_filtered, path_list)

    fill_missing_promo_numbers(partial_base, base_filtered)
    fill_missing_promo_numbers(partial_compare, compare_filtered)

    base_result = format_full_output(partial_base)
    compare_result = format_full_output(partial_compare)

    text_partial_base.delete("1.0", tk.END)
    text_partial_compare.delete("1.0", tk.END)
    text_partial_base.insert(tk.END, base_result)
    text_partial_compare.insert(tk.END, compare_result)

    highlight_promo_lines(text_partial_base)
    highlight_promo_lines(text_partial_compare)

    # ไฮไลท์ความแตกต่างใน partial JSON
    highlight_differences(text_partial_base, path_list)
    highlight_differences(text_partial_compare, path_list)

    total_diff = sum(len(diff[section]) for section in diff)
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
                path = change.path(output_format='list')
                s = "".join(f"[{p}]" if isinstance(p, int) else f"['{p}']" for p in path)
                path_list.append(s)

    partial_base = build_partial_json(base_filtered, path_list)
    partial_compare = build_partial_json(compare_filtered, path_list)

    fill_missing_promo_numbers(partial_base, base_filtered)
    fill_missing_promo_numbers(partial_compare, compare_filtered)

    wb = Workbook()
    wb.remove(wb.active)

    promos_base = partial_base.get("promoInfo", [])
    promos_compare = partial_compare.get("promoInfo", [])

    promo_dict = {}
    for promo in promos_base:
        promo_dict[promo.get("promoNumber", "N/A")] = {"base": promo, "compare": {}}
    for promo in promos_compare:
        promo_dict.setdefault(promo.get("promoNumber", "N/A"), {}).update({"compare": promo})

    highlight_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for promo_num, data in promo_dict.items():
        ws = wb.create_sheet(title=str(promo_num))
        ws.append(["Field", "Base", "Compare"])

        def recursive_write(base, compare, prefix=""):
            if isinstance(base, dict) or isinstance(compare, dict):
                all_keys = set(base.keys() if isinstance(base, dict) else []).union(compare.keys() if isinstance(compare, dict) else [])
                for key in sorted(all_keys):
                    b = base.get(key, "") if isinstance(base, dict) else ""
                    c = compare.get(key, "") if isinstance(compare, dict) else ""
                    recursive_write(b, c, f"{prefix}{key}.")
            elif isinstance(base, list) or isinstance(compare, list):
                max_len = max(len(base) if isinstance(base, list) else 0, len(compare) if isinstance(compare, list) else 0)
                for i in range(max_len):
                    b = base[i] if isinstance(base, list) and i < len(base) else ""
                    c = compare[i] if isinstance(compare, list) and i < len(compare) else ""
                    recursive_write(b, c, f"{prefix}[{i}].")
            else:
                row = [prefix.rstrip("."), base, compare]
                ws.append(row)
                if base != compare:
                    for col in range(1, 4):
                        ws.cell(row=ws.max_row, column=col).fill = highlight_fill

        recursive_write(data.get("base", {}), data.get("compare", {}))

    try:
        wb.save("JSON_Comparison_Result.xlsx")
        label_result.config(text="✅ บันทึกไฟล์ Excel แล้ว: JSON_Comparison_Result.xlsx", foreground="#66ff99")
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
    global is_fullscreen
    is_fullscreen = False
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

ttk.Label(root, text="🧠 JSON Compare Tool", style="Header.TLabel").grid(row=0, column=0, pady=(10, 5))

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

label_result = ttk.Label(root, text="", foreground="#66ff99", background=DARK_BG, font=("Segoe UI", 12, "bold"))
label_result.grid(row=3, column=0, pady=5)

frame_copy = ttk.Frame(root)
frame_copy.grid(row=4, column=0)
ttk.Button(frame_copy, text="📋 Copy Compare Diff", command=lambda: copy_text(text_partial_compare)).pack(side="left", padx=15)
ttk.Button(frame_copy, text="📋 Copy Base Diff", command=lambda: copy_text(text_partial_base)).pack(side="left", padx=15)
ttk.Button(frame_copy, text="📤 Export Excel", command=export_to_excel).pack(side="left", padx=15)  # สร้างปุ่มหลัง frame_copy สร้างแล้ว

frame_output = ttk.Frame(root)
frame_output.grid(row=5, column=0, sticky="nsew", padx=10, pady=(0,10))
frame_output.grid_columnconfigure(0, weight=1)
frame_output.grid_columnconfigure(1, weight=1)
frame_output.grid_rowconfigure(0, weight=1)

frame_diff_compare = ttk.Frame(frame_output)
frame_diff_compare.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
ttk.Label(frame_diff_compare, text="📙 JSON Compare - Differences", style="Header.TLabel").pack(anchor="w")
text_partial_compare = tk.Text(frame_diff_compare, bg=TEXTBOX_BG, fg=DARK_TEXT, insertbackground="white", relief="ridge")
text_partial_compare.pack(fill="both", expand=True)
add_right_click_menu(text_partial_compare)
bind_scroll(text_partial_compare)
bind_paste_shortcuts(text_partial_compare)

frame_diff_base = ttk.Frame(frame_output)
frame_diff_base.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
ttk.Label(frame_diff_base, text="📘 JSON Base - Differences", style="Header.TLabel").pack(anchor="w")
text_partial_base = tk.Text(frame_diff_base, bg=TEXTBOX_BG, fg=DARK_TEXT, insertbackground="white", relief="ridge")
text_partial_base.pack(fill="both", expand=True)
add_right_click_menu(text_partial_base)
bind_scroll(text_partial_base)
bind_paste_shortcuts(text_partial_base)

root.mainloop()
