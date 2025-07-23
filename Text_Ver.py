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

# ----------------- Global Variables -----------------
EXPORT_FOLDER = os.path.join(os.getcwd(), "export")
EXCEL_PATH = os.path.join(EXPORT_FOLDER, "Compare_Export.xlsx")

# ตรวจสอบและสร้างโฟลเดอร์ export หากยังไม่มี
if not os.path.exists(EXPORT_FOLDER):
    try:
        os.makedirs(EXPORT_FOLDER)
    except Exception as e:
        messagebox.showerror("ไม่สามารถสร้างโฟลเดอร์", f"เกิดข้อผิดพลาดในการสร้างโฟลเดอร์:\n{e}")
        raise

last_export_data = None  # เก็บผลลัพธ์ล่าสุดจาก compare_json

# ----------------- Excel Export Utility -----------------
def flatten_json(obj, prefix=""):
    result = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            full_key = f"{prefix}.{k}" if prefix else k
            if isinstance(v, (dict, list)):
                result.update(flatten_json(v, full_key))
            else:
                result[full_key] = v
    elif isinstance(obj, list):
        for i, item in enumerate(obj):
            full_key = f"{prefix}[{i}]"
            if isinstance(item, (dict, list)):
                result.update(flatten_json(item, full_key))
            else:
                result[full_key] = item
    return result

def export_to_excel():
    """
    ส่งออกผลการเปรียบเทียบเป็นไฟล์ Excel:
      • แต่ละ promoInfo (dict หนึ่งตัวใน list) = 1 แถว (JSON block indent 2)
      • ฝั่งที่ไม่มีข้อมูลในแถวนั้น เว้นว่าง
      • ไฮไลต์แถวที่ต่างกันเป็นสีเหลือง
      • คีย์อื่น ๆ นอก promoInfo เก็บไว้ในชีต “Others”
      • ข้อมูลในคอลัมน์ A คือ compare_online (ฝั่งซ้าย)
      • ข้อมูลในคอลัมน์ C คือ compare_newpro (ฝั่งขวา)
    """
    if not last_export_data:
        messagebox.showwarning("ยังไม่มีข้อมูลเปรียบเทียบ",
                               "กรุณาเปรียบเทียบ JSON ก่อนกด Export")
        return

    base_part, comp_part = last_export_data
    diff_fill = PatternFill(start_color="FFFF00",
                            end_color="FFFF00",
                            fill_type="solid")

    # --- เตรียมสมุดงาน ------------------------------------------------------
    try:
        if os.path.exists(EXCEL_PATH):
            wb = load_workbook(EXCEL_PATH)
        else:
            wb = Workbook()
            wb.remove(wb.active)          # ลบชีต default
    except Exception as e:
        messagebox.showerror("ไม่สามารถโหลดไฟล์ Excel", str(e))
        return

    # --- ฟังก์ชันเขียนข้อมูลลงหนึ่งชีต --------------------------------------
    def write_sheet(ws, base_data: dict, comp_data: dict):
        """
        - compare_online (comp_data) อยู่ Column A (ซ้าย)
        - compare_newpro (base_data) อยู่ Column C (ขวา)
        - แสดง JSON เป็น block (key + value ในบรรทัดเดียว)
        - ไม่แสดง null
        - ถ้ามีเฉพาะฝั่งใดฝั่งหนึ่ง → อีกฝั่งเว้นว่าง
        - ไฮไลต์เฉพาะบรรทัดที่แตกต่าง
        """

        def json_block(data):
            """แปลง dict หรือ list เป็น JSON block และไม่แสดง null"""
            if data is None:
                return []
            elif isinstance(data, (str, int, float, bool)):
                return [json.dumps(data, ensure_ascii=False)]
            elif isinstance(data, (dict, list)):
                lines = json.dumps(data, indent=2, ensure_ascii=False).splitlines()
                return [line for line in lines if line.strip() != "null"]
            else:
                return [str(data)]

        # ---------- promoInfo ----------
        if "promoInfo" in base_data or "promoInfo" in comp_data:
            base_promos = base_data.get("promoInfo", [])
            comp_promos = comp_data.get("promoInfo", [])

            max_len = max(len(base_promos), len(comp_promos))
            ws.append(["compare_online (promoInfo)", "", "compare_newpro (promoInfo)"])

            for i in range(max_len):
                base_obj = base_promos[i] if i < len(base_promos) else None
                comp_obj = comp_promos[i] if i < len(comp_promos) else None

                # สลับ base_lines กับ comp_lines เพื่อสลับฝั่งแสดงข้อมูล
                comp_lines = json_block(base_obj)  # compare_online (ซ้าย)
                base_lines = json_block(comp_obj)  # compare_newpro (ขวา)
                max_lines = max(len(base_lines), len(comp_lines))

                promo_number = None
                if base_obj and isinstance(base_obj, dict) and "promoNumber" in base_obj:
                    promo_number = base_obj["promoNumber"]
                elif comp_obj and isinstance(comp_obj, dict) and "promoNumber" in comp_obj:
                    promo_number = comp_obj["promoNumber"]

                if promo_number is not None:
                    ws.append([f"promoNumber: {promo_number}", "", f"promoNumber: {promo_number}"])

                for j in range(max_lines):
                    b_line = base_lines[j] if j < len(base_lines) else ""
                    c_line = comp_lines[j] if j < len(comp_lines) else ""

                    row_data = ["", "", ""]
                    if b_line and not c_line:
                        row_data[2] = b_line  # compare_newpro ขวา
                    elif c_line and not b_line:
                        row_data[0] = c_line  # compare_online ซ้าย
                    elif b_line and c_line:
                        row_data[2] = b_line
                        row_data[0] = c_line

                    ws.append(row_data)
                    row = ws.max_row
                    if b_line != c_line:
                        if b_line:
                            ws.cell(row=row, column=3).fill = diff_fill
                        if c_line:
                            ws.cell(row=row, column=1).fill = diff_fill

                ws.append([])

        # ---------- top-level keys ----------
        other_keys = sorted(set(base_data.keys()) | set(comp_data.keys()) - {"promoInfo"})
        if other_keys:
            ws.append([])
            ws.append(["compare_newpro", "", "compare_online"])

        for key in other_keys:
            base_val = base_data.get(key)
            comp_val = comp_data.get(key)

            base_lines = json_block(base_val)
            comp_lines = json_block(comp_val)
            max_lines = max(len(base_lines), len(comp_lines))

            if max_lines == 1:
                b_line = f'"{key}": {base_lines[0]}' if base_lines else ""
                c_line = f'"{key}": {comp_lines[0]}' if comp_lines else ""

                row_data = ["", "", ""]
                if b_line and not c_line:
                    row_data[2] = b_line
                elif c_line and not b_line:
                    row_data[0] = c_line
                elif b_line and c_line:
                    row_data[2] = b_line
                    row_data[0] = c_line

                ws.append(row_data)
                row = ws.max_row
                if b_line != c_line:
                    if b_line:
                        ws.cell(row=row, column=3).fill = diff_fill
                    if c_line:
                        ws.cell(row=row, column=1).fill = diff_fill

            else:
                ws.append([f'"{key}":', "", f'"{key}":'])
                for j in range(max_lines):
                    b_line = base_lines[j] if j < len(base_lines) else ""
                    c_line = comp_lines[j] if j < len(comp_lines) else ""

                    row_data = ["", "", ""]
                    if b_line and not c_line:
                        row_data[2] = b_line
                    elif c_line and not b_line:
                        row_data[0] = c_line
                    elif b_line and c_line:
                        row_data[2] = b_line
                        row_data[0] = c_line

                    ws.append(row_data)
                    row = ws.max_row
                    if b_line != c_line:
                        if b_line:
                            ws.cell(row=row, column=3).fill = diff_fill
                        if c_line:
                            ws.cell(row=row, column=1).fill = diff_fill

        ws.column_dimensions["A"].width = 60
        ws.column_dimensions["B"].width = 5
        ws.column_dimensions["C"].width = 60


   # --- สร้างชีตเดียวรวมทุก promoNumber --------------------------------------
    if "AllPromos" in wb.sheetnames:
        del wb["AllPromos"]
    ws = wb.create_sheet("AllPromos")

    base_promos = {p["promoNumber"]: p
                for p in base_part.get("promoInfo", [])
                if "promoNumber" in p}
    comp_promos = {p["promoNumber"]: p
                for p in comp_part.get("promoInfo", [])
                if "promoNumber" in p}

    all_promo_nums = sorted(set(base_promos) | set(comp_promos),
                            key=lambda x: int(x) if str(x).isdigit() else str(x))

    for promo_num in all_promo_nums:
        ws.append([f"========== promoNumber: {promo_num} ==========", "", ""])
        write_sheet(ws,
                    base_promos.get(promo_num, {}),
                    comp_promos.get(promo_num, {}))

    # --- เพิ่มข้อมูล Others ต่อท้ายในชีตเดียวกัน -----------------------------
    others_base = {k: v for k, v in base_part.items() if k != "promoInfo"}
    others_comp = {k: v for k, v in comp_part.items() if k != "promoInfo"}

    if others_base or others_comp:
        ws.append([])
        ws.append(["========== Others ==========", "", ""])
        write_sheet(ws, others_base, others_comp)



    # --- บันทึก -------------------------------------------------------------
    try:
        wb.save(EXCEL_PATH)
        messagebox.showinfo("ส่งออกสำเร็จ",
                            f"ไฟล์ Excel ถูกบันทึกที่:\n{EXCEL_PATH}")
    except PermissionError:
        messagebox.showerror("ไม่สามารถบันทึก",
                             "กรุณาปิดไฟล์ Excel ก่อนแล้วลองใหม่")


# ----------------- Core Function: compare_json -----------------

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

    base_promos = {p["promoNumber"]: p for p in base_filtered.get("promoInfo", []) if "promoNumber" in p}
    compare_promos = {p["promoNumber"]: p for p in compare_filtered.get("promoInfo", []) if "promoNumber" in p}

    partial_base_result = {"promoInfo": []}
    partial_compare_result = {"promoInfo": []}
    total_diff_paths = []

    all_promo_numbers = sorted(set(base_promos.keys()) | set(compare_promos.keys()), key=lambda x: int(x))

    for promo_num in all_promo_numbers:
        base_promo = base_promos.get(promo_num)
        compare_promo = compare_promos.get(promo_num)

        if base_promo and compare_promo:
            # กรณีมีทั้งสองฝั่ง
            diff = DeepDiff(base_promo, compare_promo, ignore_order=False, report_repetition=True, view="tree")

            if not diff:
                partial_base = base_promo.copy()
                partial_base["promoNumber"] = promo_num
                partial_compare = compare_promo.copy()
                partial_compare["promoNumber"] = promo_num

            path_list = []
            for section in diff:
                for change in diff[section]:
                    if hasattr(change, 'path'):
                        path = change.path(output_format='list')
                        s = "".join(f"[{p}]" if isinstance(p, int) else f"['{p}']" for p in path)
                        path_list.append(s)

            total_diff_paths.extend([f"['promoInfo'][{len(partial_base_result['promoInfo'])}]{p}" for p in path_list])

            partial_base = build_partial_json(base_promo, path_list)
            partial_base["promoNumber"] = promo_num
            partial_compare = build_partial_json(compare_promo, path_list)
            partial_compare["promoNumber"] = promo_num

        elif base_promo and not compare_promo:
            # มีเฉพาะใน Base
            partial_base = base_promo.copy()
            partial_base["promoNumber"] = promo_num
            partial_compare = {"promoNumber": promo_num}  # ว่างเปล่า

        elif compare_promo and not base_promo:
            # มีเฉพาะใน Compare
            partial_compare = compare_promo.copy()
            partial_compare["promoNumber"] = promo_num
            partial_base = {"promoNumber": promo_num}  # ว่างเปล่า

        partial_base_result["promoInfo"].append(partial_base)
        partial_compare_result["promoInfo"].append(partial_compare)

    # ==== เปรียบเทียบฟิลด์อื่น ๆ ที่ไม่ใช่ promoInfo ====
    other_keys = set(base_filtered.keys()) | set(compare_filtered.keys())
    other_keys.discard("promoInfo")

    for key in sorted(other_keys):
        if key not in base_filtered or key not in compare_filtered:
            continue

        diff = DeepDiff(base_filtered[key], compare_filtered[key], ignore_order=False, report_repetition=True, view="tree")

        if not diff:
            continue

        path_list = []
        for section in diff:
            for change in diff[section]:
                if hasattr(change, 'path'):
                    path = change.path(output_format='list')
                    s = f"['{key}']" + "".join(f"[{p}]" if isinstance(p, int) else f"['{p}']" for p in path)
                    path_list.append(s)

        total_diff_paths.extend(path_list)

        partial_base = build_partial_json(base_filtered, path_list)
        partial_compare = build_partial_json(compare_filtered, path_list)

        partial_base_result.update(partial_base)
        partial_compare_result.update(partial_compare)

    # ==== สร้างผลลัพธ์และแสดงผล ====
    base_result = format_full_output(partial_base_result)
    compare_result = format_full_output(partial_compare_result)

    text_partial_base.delete("1.0", tk.END)
    text_partial_compare.delete("1.0", tk.END)
    text_partial_base.insert(tk.END, base_result)
    text_partial_compare.insert(tk.END, compare_result)

    highlight_promo_lines(text_partial_base)
    highlight_promo_lines(text_partial_compare)
    highlight_differences(text_partial_base, total_diff_paths)
    highlight_differences(text_partial_compare, total_diff_paths)

    global last_export_data
    last_export_data = (partial_base_result, partial_compare_result)

    label_result.config(text=f"🔍 พบความแตกต่างทั้งหมด {len(total_diff_paths)} จุด")
 

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
