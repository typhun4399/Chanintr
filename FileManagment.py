import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
from collections import defaultdict, Counter
import threading
from functools import partial
import pandas as pd

# We need a third-party library to move files to the recycle bin safely.
# The user needs to install it by running: pip install send2trash
try:
    import send2trash
except ImportError:
    messagebox.showerror("Missing Library", "Could not find the 'send2trash' library.\nPlease install it by running:\n\npip install send2trash")
    exit()

# ==============================================================================
# SECTION 1: CORE LOGIC
# ==============================================================================

# --- Reporting Tasks (1-4) ---

def run_task_list_contents(base_path, output_file, update_status, selected_subfolders):
    update_status("Task 1: กำลังรวบรวมข้อมูลไฟล์...")
    results = []
    all_folders = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    for i, folder in enumerate(all_folders):
        folder_path = os.path.join(base_path, folder)
        update_status(f"Task 1: กำลังตรวจสอบ {folder} ({i+1}/{len(all_folders)})")
        folder_data = {}
        for sub_folder_name in selected_subfolders:
            sub_path = os.path.join(folder_path, sub_folder_name)
            if os.path.isdir(sub_path):
                files = [f for f in os.listdir(sub_path) if os.path.isfile(os.path.join(sub_path, f))]
                folder_data[sub_folder_name] = files if files else None
        if folder_data:
            results.append((folder, folder_data))
    update_status("Task 1: กำลังเขียนไฟล์รายงาน...")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(f"📊 รายงานไฟล์ในโฟลเดอร์ย่อย: {', '.join(selected_subfolders)}\n\n")
        for folder, data in sorted(results):
            f.write(f"📁 {folder}\n")
            for key in selected_subfolders:
                if key in data:
                    files = data[key]
                    if files:
                        f.write(f"  - {key}: {len(files)} ไฟล์\n")
                        for file in sorted(files):
                            f.write(f"    - {file}\n")
                    else:
                        f.write(f"  - {key}: 0 ไฟล์ (โฟลเดอร์ว่าง)\n")
            f.write("\n" + "="*50 + "\n\n")

def run_task_find_empty_subfolders(base_path, output_file, update_status, selected_subfolders):
    update_status("Task 2: กำลังค้นหาโฟลเดอร์ย่อยที่ว่างเปล่า...")
    empty_folders = []
    all_folders = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    for i, folder in enumerate(all_folders):
        folder_path = os.path.join(base_path, folder)
        update_status(f"Task 2: กำลังตรวจสอบ {folder} ({i+1}/{len(all_folders)})")
        try:
            for sub in os.listdir(folder_path):
                if sub in selected_subfolders:
                    sub_path = os.path.join(folder_path, sub)
                    if os.path.isdir(sub_path) and not os.listdir(sub_path):
                        empty_folders.append((folder, sub))
        except OSError as e:
            print(f"ไม่สามารถเข้าถึง {folder_path}: {e}")
    update_status("Task 2: กำลังเขียนไฟล์รายงาน...")
    with open(output_file, "w", encoding="utf-8") as f:
        if not empty_folders:
            f.write(f"✅ ไม่พบโฟลเดอร์ย่อย ({', '.join(selected_subfolders)}) ที่ว่างเปล่า\n")
        else:
            f.write(f"⚠️ รายงานโฟลเดอร์ย่อยที่ว่างเปล่า (จากตัวเลือก: {', '.join(selected_subfolders)}):\n\n")
            for folder, sub in sorted(empty_folders):
                f.write(f"{folder}\\{sub}\n")

def run_task_summarize_types(base_path, output_file, update_status, selected_subfolders):
    update_status("Task 3: กำลังนับประเภทไฟล์...")
    summary = {name: Counter() for name in selected_subfolders}
    all_folders = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    for i, folder in enumerate(all_folders):
        folder_path = os.path.join(base_path, folder)
        update_status(f"Task 3: กำลังตรวจสอบ {folder} ({i+1}/{len(all_folders)})")
        for sub in selected_subfolders:
            sub_path = os.path.join(folder_path, sub)
            if os.path.isdir(sub_path):
                for root, dirs, files in os.walk(sub_path):
                    for file in files:
                        ext = os.path.splitext(file)[1].lower()
                        if ext:
                            summary[sub][ext] += 1
    update_status("Task 3: กำลังเขียนไฟล์รายงาน...")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(f"📊 สรุปประเภทและจำนวนไฟล์ทั้งหมด (จากตัวเลือก: {', '.join(selected_subfolders)})\n")
        f.write("="*50 + "\n\n")
        total_files = 0
        for sub in selected_subfolders:
            f.write(f"📁 หมวดหมู่: {sub}\n")
            if not summary.get(sub):
                f.write("  (ไม่พบไฟล์)\n\n")
            else:
                sub_total = sum(summary[sub].values())
                total_files += sub_total
                for ext, count in sorted(summary[sub].items()):
                    f.write(f"  - {ext:<10} : {count} ไฟล์\n")
                f.write(f"  -- รวม {sub_total} ไฟล์ --\n\n")
        f.write(f"\n\n✨ รวมไฟล์ทั้งหมดที่พบ: {total_files} ไฟล์")

def run_task_find_specific_files(base_path, output_file, update_status, extensions, search_scope, selected_subfolders=None):
    update_status(f"Task 4: กำลังค้นหาไฟล์: {', '.join(extensions)}...")
    results = []
    for root, dirs, files in os.walk(base_path):
        is_in_scope = False
        if search_scope == 'all':
            is_in_scope = True
        else: 
            if selected_subfolders:
                rel_path_parts = os.path.relpath(root, base_path).split(os.path.sep)
                if any(part in selected_subfolders for part in rel_path_parts):
                    is_in_scope = True
        if is_in_scope:
            for file in files:
                if any(file.lower().endswith(ext) for ext in extensions):
                    path_parts = root.replace(base_path, '').strip(os.path.sep).split(os.path.sep)
                    main_folder = path_parts[0] if path_parts else os.path.basename(root)
                    relative_path = os.path.relpath(root, base_path)
                    results.append((main_folder, relative_path, file))
    update_status("Task 4: กำลังเขียนไฟล์รายงาน...")
    with open(output_file, "w", encoding="utf-8") as f:
        if not results:
            f.write(f"✅ ไม่พบไฟล์ที่ตรงกับนามสกุล ({', '.join(extensions)}) ในขอบเขตที่เลือก\n")
        else:
            f.write(f"🔹 รายงานการค้นหาไฟล์: {', '.join(extensions)}\n" + "="*50 + "\n\n")
            grouped_results = defaultdict(list)
            for main_folder, rel_path, filename in sorted(results):
                grouped_results[main_folder].append((rel_path, filename))
            for main_folder, items in grouped_results.items():
                f.write(f"📁 โฟลเดอร์หลัก: {main_folder}\n")
                for rel_path, filename in items:
                    f.write(f"  - พบ '{filename}' ใน: .\\{rel_path}\n")
                f.write("\n")

# --- Deletion Tasks (5-7) ---

def run_task_delete_files_by_ext(base_path, output_file, update_status, selected_subfolders, extensions):
    update_status(f"Task 5: กำลังค้นหาและย้ายไฟล์: {', '.join(extensions)}...")
    deleted_log = []
    all_folders = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    for i, folder in enumerate(all_folders):
        folder_path = os.path.join(base_path, folder)
        update_status(f"Task 5: กำลังสแกนใน {folder} ({i+1}/{len(all_folders)})")
        for sub in selected_subfolders:
            sub_path = os.path.join(folder_path, sub)
            if os.path.isdir(sub_path):
                for root, _, files in os.walk(sub_path):
                    for file in files:
                        if any(file.lower().endswith(ext) for ext in extensions):
                            file_path = os.path.join(root, file)
                            try:
                                send2trash.send2trash(file_path)
                                deleted_log.append(file_path)
                            except Exception as e:
                                print(f"ย้ายไม่ได้: {file_path} ({e})")
                                deleted_log.append(f"ย้ายไม่สำเร็จ: {file_path} (Error: {e})")
    
    update_status("Task 5: กำลังเขียนไฟล์รายงาน...")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(f"🗑️ รายงานการย้ายไฟล์นามสกุล {', '.join(extensions)} ไปที่ถังขยะ:\n")
        f.write(f"(ตรวจสอบใน: {', '.join(selected_subfolders)})\n\n")
        if not deleted_log:
            f.write("✅ ไม่พบไฟล์ที่ตรงเงื่อนไข\n")
        else:
            for log_entry in deleted_log:
                f.write(f"- {log_entry}\n")

def run_task_delete_empty_subfolders(base_path, output_file, update_status, scope, selected_subfolders):
    update_status("Task 6: กำลังค้นหาและย้ายโฟลเดอร์ย่อยที่ว่างเปล่า...")
    deleted_log = []
    all_folders = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    for i, folder in enumerate(all_folders):
        folder_path = os.path.join(base_path, folder)
        update_status(f"Task 6: กำลังตรวจสอบ {folder} ({i+1}/{len(all_folders)})")
        
        subfolders_to_check = []
        if scope == 'all':
            subfolders_to_check = [d for d in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, d))]
        else: # scope == 'selected'
            subfolders_to_check = selected_subfolders

        for sub in subfolders_to_check:
            sub_path = os.path.join(folder_path, sub)
            if os.path.isdir(sub_path) and not os.listdir(sub_path):
                try:
                    send2trash.send2trash(sub_path)
                    deleted_log.append(sub_path)
                except Exception as e:
                    print(f"ย้ายไม่ได้: {sub_path} ({e})")
                    deleted_log.append(f"ย้ายไม่สำเร็จ: {sub_path} (Error: {e})")

    update_status("Task 6: กำลังเขียนไฟล์รายงาน...")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(f"🗑️ รายงานการย้ายโฟลเดอร์ย่อยที่ว่างเปล่าไปที่ถังขยะ:\n")
        f.write(f"(ขอบเขต: {'ทุกโฟลเดอร์ย่อย' if scope == 'all' else 'เฉพาะที่เลือก'})\n\n")
        if not deleted_log:
            f.write("✅ ไม่พบโฟลเดอร์ย่อยที่ว่างเปล่าให้ย้าย\n")
        else:
            for log_entry in deleted_log:
                f.write(f"- {log_entry}\n")

def run_task_delete_empty_main_folders(base_path, output_file, update_status, selected_subfolders):
    update_status("Task 7: กำลังค้นหาโฟลเดอร์หลักที่ว่างเปล่า...")
    deleted_log = []
    all_folders = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    for i, folder_name in enumerate(all_folders):
        folder_path = os.path.join(base_path, folder_name)
        update_status(f"Task 7: กำลังตรวจสอบ {folder_name} ({i+1}/{len(all_folders)})")
        
        has_any_files = False
        for sub_folder_name in selected_subfolders:
            sub_path = os.path.join(folder_path, sub_folder_name)
            if os.path.isdir(sub_path):
                if any(os.path.isfile(os.path.join(sub_path, f)) for f in os.listdir(sub_path)):
                    has_any_files = True
                    break
        
        if not has_any_files:
            try:
                send2trash.send2trash(folder_path)
                deleted_log.append(folder_path)
                update_status(f"Task 7: 🗑️ ย้ายแล้ว: {folder_name}")
            except Exception as e:
                print(f"ไม่สามารถย้ายโฟลเดอร์ '{folder_name}': {e}")
                deleted_log.append(f"ย้ายไม่สำเร็จ: {folder_path} (Error: {e})")

    update_status("Task 7: กำลังเขียนไฟล์รายงาน...")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("🗑️ รายงานการย้ายโฟลเดอร์หลัก (ที่ไม่มีไฟล์ในโฟลเดอร์ย่อยที่เลือก) ไปที่ถังขยะ:\n")
        f.write(f"(ตรวจสอบใน: {', '.join(selected_subfolders)})\n\n")
        if not deleted_log:
            f.write("✅ ไม่พบโฟลเดอร์หลักที่ตรงเงื่อนไข\n")
        else:
            for log_entry in deleted_log:
                f.write(f"- {log_entry}\n")

# --- Update Task (8) ---

def run_task_update_excel(base_path, output_excel_path, update_status, excel_input_path, selected_subfolders_for_notes, id_col, status_col, notes_col):
    update_status("Task 8: Starting Excel update process...")
    try:
        df = pd.read_excel(excel_input_path)
        update_status(f"Task 8: Successfully read '{os.path.basename(excel_input_path)}'")
    except Exception as e:
        messagebox.showerror("Excel Read Error", f"Could not read the Excel file:\n{e}")
        update_status(f"Error reading Excel file.")
        return

    required_cols = [id_col, status_col, notes_col]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        messagebox.showerror("Missing Columns", f"The Excel file is missing required columns: {', '.join(missing_cols)}")
        update_status("Excel file missing columns.")
        return

    df[status_col] = "N"
    df[notes_col] = ""
    update_status("Task 8: Reset status and notes columns.")

    try:
        subdirectories = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    except Exception as e:
        messagebox.showerror("Folder Error", f"Could not read directories in the target folder:\n{e}")
        update_status("Error reading target folder.")
        return

    def has_files(path):
        return os.path.exists(path) and any(os.path.isfile(os.path.join(path, f)) for f in os.listdir(path))

    total_rows = len(df)
    for i, row in df.iterrows():
        if pd.isna(row[id_col]):
            continue
        folder_id = str(row[id_col]).strip()
        update_status(f"Task 8: Processing ID {folder_id} ({i+1}/{total_rows})")
        
        matched_folders = [f for f in subdirectories if f.split("_")[0] == folder_id]

        if matched_folders:
            found_any_content_at_all = False
            for sub_folder in matched_folders:
                folder_path = os.path.join(base_path, sub_folder)
                try:
                    for item in os.listdir(folder_path):
                        item_path = os.path.join(folder_path, item)
                        if os.path.isdir(item_path) and has_files(item_path):
                            found_any_content_at_all = True
                            break
                except OSError:
                    continue
                if found_any_content_at_all:
                    break

            notes_list = []
            for sub_folder in matched_folders:
                folder_path = os.path.join(base_path, sub_folder)
                for name in selected_subfolders_for_notes:
                    if not has_files(os.path.join(folder_path, name)):
                        notes_list.append(f"ไม่มี{name}")

            if found_any_content_at_all:
                df.at[i, status_col] = "Y"
                df.at[i, notes_col] = "/".join(sorted(list(set(notes_list))))
            else:
                df.at[i, status_col] = "N"
                df.at[i, notes_col] = "/".join(sorted([f"ไม่มี{name}" for name in selected_subfolders_for_notes]))
        else:
            df.at[i, status_col] = "N"
            df.at[i, notes_col] = "/".join(sorted([f"ไม่มี{name}" for name in selected_subfolders_for_notes]))
    try:
        df.to_excel(output_excel_path, index=False)
        update_status(f"Task 8: Successfully saved updated file to '{os.path.basename(output_excel_path)}'")
    except Exception as e:
        messagebox.showerror("Excel Save Error", f"Could not save the updated Excel file:\n{e}")
        update_status("Error saving Excel file.")

# ==============================================================================
# SECTION 2: TKINTER GUI APPLICATION
# ==============================================================================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("File Management Toolkit")
        try:
            self.root.state('zoomed')
        except tk.TclError:
            self.root.attributes('-zoomed', True)
        
        self.setup_styles()

        self.base_path = tk.StringVar()
        self.output_file = tk.StringVar()
        self.selected_task = tk.IntVar(value=1)
        self.subfolder_vars = {}
        self.extensions_to_find = tk.StringVar(value=".pdf, .dwg, .step")
        self.extensions_to_delete = tk.StringVar(value=".tmp, .crdownload, .log")
        self.search_scope = tk.StringVar(value="all")
        self.delete_empty_scope = tk.StringVar(value="selected")
        self.excel_input_path = tk.StringVar()
        self.id_col = tk.StringVar()
        self.status_col = tk.StringVar()
        self.notes_col = tk.StringVar()

        self.create_widgets()

    def setup_styles(self):
        style = ttk.Style(self.root)
        self.default_font = font.Font(family="Tahoma", size=10)
        self.title_font = font.Font(family="Tahoma", size=11, weight="bold")
        self.warning_font = font.Font(family="Tahoma", size=10, weight="bold")

        style.configure('.', font=self.default_font)
        style.configure('TLabel', padding=2)
        style.configure('TButton', padding=5)
        style.configure('TRadiobutton', padding=2)
        style.configure('TLabelframe.Label', font=self.title_font)
        style.configure('Warning.TLabel', foreground='red', font=self.warning_font)
        style.configure('TNotebook.Tab', font=self.default_font, padding=[10, 5])

    def create_widgets(self):
        # NEW: Main scrollable canvas setup
        main_canvas_container = ttk.Frame(self.root)
        main_canvas_container.pack(fill=tk.BOTH, expand=True)
        self.main_canvas = tk.Canvas(main_canvas_container, borderwidth=0, highlightthickness=0)
        main_scrollbar = ttk.Scrollbar(main_canvas_container, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=main_scrollbar.set)
        main_scrollbar.pack(side="right", fill="y")
        self.main_canvas.pack(side="left", fill="both", expand=True)

        self.content_frame = ttk.Frame(self.main_canvas, padding="10")
        self.content_frame_id = self.main_canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
        
        self.content_frame.bind("<Configure>", lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.main_canvas.bind("<Configure>", self._on_main_canvas_configure)
        
        # This is the crucial part for global scrolling
        self.root.bind_all("<MouseWheel>", self._on_global_mousewheel)
        self.root.bind_all("<Button-4>", self._on_global_mousewheel)
        self.root.bind_all("<Button-5>", self._on_global_mousewheel)

        path_frame = ttk.LabelFrame(self.content_frame, text="1. เลือกโฟลเดอร์และไฟล์", padding="10")
        path_frame.pack(fill=tk.X, pady=5)
        ttk.Label(path_frame, text="โฟลเดอร์เป้าหมาย:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.base_path_entry = ttk.Entry(path_frame, textvariable=self.base_path, width=60)
        self.base_path_entry.grid(row=0, column=1, sticky=tk.EW, padx=5)
        ttk.Button(path_frame, text="Browse...", command=self.browse_base_path).grid(row=0, column=2)
        ttk.Label(path_frame, text="บันทึกไฟล์ Output เป็น:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.output_file_entry = ttk.Entry(path_frame, textvariable=self.output_file, width=60)
        self.output_file_entry.grid(row=1, column=1, sticky=tk.EW, padx=5)
        ttk.Button(path_frame, text="Browse...", command=self.browse_output_file).grid(row=1, column=2)
        path_frame.columnconfigure(1, weight=1)

        subfolder_frame = ttk.LabelFrame(self.content_frame, text="2. เลือกโฟลเดอร์ย่อย (สำหรับทุกโหมด)", padding="10")
        subfolder_frame.pack(fill=tk.X, pady=10)
        self.scan_button = ttk.Button(subfolder_frame, text="Scan for Subfolders", command=self.scan_subfolders, state=tk.DISABLED)
        self.scan_button.pack(pady=5)
        self.subfolder_canvas, self.subfolder_checkbox_frame = self._create_scrollable_frame(subfolder_frame, 100)
        
        notebook = ttk.Notebook(self.content_frame)
        notebook.pack(fill=tk.BOTH, pady=5, expand=True)
        reporting_frame = ttk.Frame(notebook, padding="10")
        deletion_frame = ttk.Frame(notebook, padding="10")
        update_frame = ttk.Frame(notebook, padding="10")
        
        notebook.add(reporting_frame, text='📊 รายงาน (Reporting)')
        notebook.add(deletion_frame, text='🗑️ ย้ายไปที่ถังขยะ (Deletion)')
        notebook.add(update_frame, text='🔄 อัปเดต Excel (Update)')
        notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)


        self.create_reporting_tab(reporting_frame)
        self.create_deletion_tab(deletion_frame)
        self.create_update_tab(update_frame)
        
        control_frame = ttk.Frame(self.content_frame, padding="10")
        control_frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_label = ttk.Label(control_frame, text="สถานะ: พร้อมทำงาน", anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.run_button = ttk.Button(control_frame, text="เริ่มทำงาน", command=self.start_task)
        self.run_button.pack(side=tk.RIGHT)
        
        self.toggle_ui_elements()

    def _create_scrollable_frame(self, parent, max_height):
        scroll_container = ttk.Frame(parent)
        scroll_container.pack(fill=tk.X, expand=True)
        canvas = tk.Canvas(scroll_container, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        checkbox_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=checkbox_frame, anchor="nw")
        
        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            self.root.update_idletasks()
            req_height = checkbox_frame.winfo_reqheight()
            canvas.config(height=min(req_height, max_height))

        checkbox_frame.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas.create_window, width=e.width))
        return canvas, checkbox_frame

    def create_reporting_tab(self, parent_frame):
        task_frame = ttk.LabelFrame(parent_frame, text="3. เลือกการทำงาน", padding="10")
        task_frame.pack(fill=tk.BOTH, pady=5, expand=True)
        tasks = [("รายงานรายละเอียดไฟล์", 1), ("ค้นหาโฟลเดอร์ย่อยที่ว่างเปล่า", 2), ("สรุปประเภทไฟล์", 3)]
        for text, value in tasks:
            ttk.Radiobutton(task_frame, text=f"{text} (ตามที่เลือก)", variable=self.selected_task, value=value, command=self.toggle_ui_elements).pack(anchor=tk.W)
        
        task_4_container = ttk.Frame(task_frame)
        task_4_container.pack(fill=tk.X, anchor=tk.W, pady=5)
        ttk.Radiobutton(task_4_container, text="ค้นหาไฟล์ตามนามสกุล", variable=self.selected_task, value=4, command=self.toggle_ui_elements).pack(side=tk.LEFT, anchor=tk.NW, pady=5)
        
        self.task_4_options_frame = ttk.Frame(task_4_container)
        self.task_4_options_frame.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        ext_frame = ttk.Frame(self.task_4_options_frame)
        ext_frame.pack(fill=tk.X)
        ttk.Label(ext_frame, text="นามสกุลไฟล์ (คั่นด้วย comma):").pack(side=tk.LEFT, anchor=tk.W)
        self.extensions_entry = ttk.Entry(ext_frame, textvariable=self.extensions_to_find, width=40)
        self.extensions_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.scope_frame = ttk.Frame(self.task_4_options_frame)
        self.scope_frame.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(self.scope_frame, text="สแกนทุกที่", variable=self.search_scope, value="all", command=self.toggle_ui_elements).pack(side=tk.LEFT)
        ttk.Radiobutton(self.scope_frame, text="สแกนในโฟลเดอร์ที่เลือก", variable=self.search_scope, value="selected", command=self.toggle_ui_elements).pack(side=tk.LEFT, padx=10)

    def create_deletion_tab(self, parent_frame):
        ttk.Label(parent_frame, text="คำเตือน: การดำเนินการในหน้านี้จะย้ายไฟล์/โฟลเดอร์ไปที่ถังขยะ", style='Warning.TLabel').pack(pady=5, fill=tk.X)
        task_frame = ttk.LabelFrame(parent_frame, text="3. เลือกประเภทการย้ายไปที่ถังขยะ", padding="10")
        task_frame.pack(fill=tk.BOTH, pady=5, expand=True)
        
        task_5_container = ttk.Frame(task_frame)
        task_5_container.pack(fill=tk.X, anchor=tk.W, pady=5)
        ttk.Radiobutton(task_5_container, text="ย้ายไฟล์ตามนามสกุล (ในโฟลเดอร์ที่เลือก)", variable=self.selected_task, value=5, command=self.toggle_ui_elements).pack(side=tk.LEFT, anchor=tk.NW, pady=5)
        self.task_5_options_frame = ttk.Frame(task_5_container)
        self.task_5_options_frame.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        ttk.Label(self.task_5_options_frame, text="นามสกุลไฟล์ที่จะย้าย:").pack(side=tk.LEFT, anchor=tk.W)
        self.delete_extensions_entry = ttk.Entry(self.task_5_options_frame, textvariable=self.extensions_to_delete, width=40)
        self.delete_extensions_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        task_6_container = ttk.Frame(task_frame)
        task_6_container.pack(fill=tk.X, anchor=tk.W, pady=5)
        ttk.Radiobutton(task_6_container, text="ย้ายโฟลเดอร์ย่อยที่ว่างเปล่า", variable=self.selected_task, value=6, command=self.toggle_ui_elements).pack(side=tk.LEFT, anchor=tk.NW, pady=5)
        self.task_6_scope_frame = ttk.Frame(task_6_container)
        self.task_6_scope_frame.pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(self.task_6_scope_frame, text="เฉพาะโฟลเดอร์ย่อยที่เลือก", variable=self.delete_empty_scope, value="selected", command=self.toggle_ui_elements).pack(side=tk.LEFT)
        ttk.Radiobutton(self.task_6_scope_frame, text="ทุกโฟลเดอร์ย่อยที่พบ", variable=self.delete_empty_scope, value="all", command=self.toggle_ui_elements).pack(side=tk.LEFT, padx=10)
        
        ttk.Radiobutton(task_frame, text="ย้ายโฟลเดอร์หลักที่ว่างเปล่า (ตามที่เลือก)", variable=self.selected_task, value=7, command=self.toggle_ui_elements).pack(anchor=tk.W, pady=5)

    def create_update_tab(self, parent_frame):
        # NEW: Removed the Radiobutton for task 8
        ttk.Label(parent_frame, text="อัปเดตสถานะในไฟล์ Excel ตามโฟลเดอร์ที่เลือกใน Section 2", wraplength=500).pack(anchor=tk.W, pady=5)
        
        excel_frame = ttk.LabelFrame(parent_frame, text="เลือกไฟล์ Excel", padding="10")
        excel_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(excel_frame, text="ไฟล์ Excel ต้นฉบับ:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.excel_input_entry = ttk.Entry(excel_frame, textvariable=self.excel_input_path, width=60)
        self.excel_input_entry.grid(row=0, column=1, sticky=tk.EW, padx=5)
        ttk.Button(excel_frame, text="Browse...", command=self.browse_excel_input).grid(row=0, column=2)
        excel_frame.columnconfigure(1, weight=1)
        ttk.Label(excel_frame, text="(ไฟล์ Output จะใช้เส้นทางจากข้อ 1)").grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        column_map_frame = ttk.LabelFrame(parent_frame, text="กำหนดคอลัมน์", padding="10")
        column_map_frame.pack(fill=tk.X, pady=5)
        self.read_columns_button = ttk.Button(column_map_frame, text="Read Columns from Excel", command=self.read_excel_columns, state=tk.DISABLED)
        self.read_columns_button.pack(pady=5)

        combo_frame = ttk.Frame(column_map_frame)
        combo_frame.pack(fill=tk.X, expand=True, pady=5)
        
        ttk.Label(combo_frame, text="ID Column:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.id_col_combo = ttk.Combobox(combo_frame, textvariable=self.id_col, state='readonly')
        self.id_col_combo.grid(row=0, column=1, sticky=tk.EW, padx=5)
        
        ttk.Label(combo_frame, text="Status Column:").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.status_col_combo = ttk.Combobox(combo_frame, textvariable=self.status_col, state='readonly')
        self.status_col_combo.grid(row=1, column=1, sticky=tk.EW, padx=5)
        
        ttk.Label(combo_frame, text="Notes Column:").grid(row=2, column=0, sticky=tk.W, padx=5)
        self.notes_col_combo = ttk.Combobox(combo_frame, textvariable=self.notes_col, state='readonly')
        self.notes_col_combo.grid(row=2, column=1, sticky=tk.EW, padx=5)
        combo_frame.columnconfigure(1, weight=1)

    def _on_main_canvas_configure(self, event):
        self.main_canvas.itemconfig(self.content_frame_id, width=event.width)

    def _on_global_mousewheel(self, event):
        # This function is now empty
        pass

    def on_tab_changed(self, event):
        try:
            selected_tab_index = event.widget.index(event.widget.select())
            if selected_tab_index == 0: # Reporting Tab
                if self.selected_task.get() not in [1, 2, 3, 4]:
                    self.selected_task.set(1)
            elif selected_tab_index == 1: # Deletion Tab
                if self.selected_task.get() not in [5, 6, 7]:
                    self.selected_task.set(5)
            elif selected_tab_index == 2: # Update Tab
                self.selected_task.set(8)
            
            self.toggle_ui_elements()
        except tk.TclError:
            pass

    def browse_base_path(self):
        path = filedialog.askdirectory(title="เลือกโฟลเดอร์ที่ต้องการตรวจสอบ")
        if path:
            self.base_path.set(path)
            self.scan_button.config(state=tk.NORMAL)
            self.update_status("กรุณากด Scan for Subfolders")
            for widget in self.subfolder_checkbox_frame.winfo_children(): widget.destroy()
            self.subfolder_vars.clear()
    def browse_output_file(self):
        task_id = self.selected_task.get()
        if task_id == 8:
            filetypes = [("Excel files", "*.xlsx"), ("All files", "*.*")]
            defaultext = ".xlsx"
        else:
            filetypes = [("Text files", "*.txt"), ("All files", "*.*")]
            defaultext = ".txt"
        path = filedialog.asksaveasfilename(title="เลือกตำแหน่งและตั้งชื่อไฟล์ Output", defaultextension=defaultext, filetypes=filetypes)
        if path: self.output_file.set(path)
    def browse_excel_input(self):
        path = filedialog.askopenfilename(title="เลือกไฟล์ Excel ต้นฉบับ", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.excel_input_path.set(path)
            self.read_columns_button.config(state=tk.NORMAL)
            self.update_status("กรุณากด Read Columns from Excel")
            self.id_col_combo.set(''); self.status_col_combo.set(''); self.notes_col_combo.set('')
            self.id_col_combo['values'] = []; self.status_col_combo['values'] = []; self.notes_col_combo['values'] = []
            if not self.output_file.get():
                output_path = os.path.splitext(path)[0] + "_updated.xlsx"
                self.output_file.set(output_path)
    
    def read_excel_columns(self):
        path = self.excel_input_path.get()
        if not path:
            messagebox.showerror("ผิดพลาด", "กรุณาเลือกไฟล์ Excel ก่อน")
            return
        try:
            self.update_status("Reading Excel columns...")
            columns = pd.read_excel(path, nrows=0).columns.tolist()
            self.id_col_combo['values'] = columns
            self.status_col_combo['values'] = columns
            self.notes_col_combo['values'] = columns
            if 'id' in columns: self.id_col_combo.set('id')
            if '2D/ 3D/ Datasheet' in columns: self.status_col_combo.set('2D/ 3D/ Datasheet')
            if 'Notes' in columns: self.notes_col_combo.set('Notes')
            self.update_status("Read columns successfully. Please map them.")
        except Exception as e:
            messagebox.showerror("Excel Read Error", f"Could not read columns from the Excel file:\n{e}")
            self.update_status("Failed to read Excel columns.")

    def update_status(self, message):
        self.status_label.config(text=f"สถานะ: {message}")
        self.root.update_idletasks()
    def scan_subfolders(self):
        base_path = self.base_path.get()
        if not base_path or not os.path.isdir(base_path): messagebox.showerror("ผิดพลาด", "กรุณาเลือกโฟลเดอร์เป้าหมายที่ถูกต้องก่อน"); return
        self.update_status("กำลังสแกนหาโฟลเดอร์ย่อย..."); self.scan_button.config(state=tk.DISABLED)
        for widget in self.subfolder_checkbox_frame.winfo_children(): widget.destroy()
        self.subfolder_vars.clear()
        unique_subfolders = set()
        try:
            for item in os.listdir(base_path):
                item_path = os.path.join(base_path, item)
                if os.path.isdir(item_path):
                    for sub_item in os.listdir(item_path):
                        if os.path.isdir(os.path.join(item_path, sub_item)): unique_subfolders.add(sub_item)
        except Exception as e:
            messagebox.showerror("Error Scanning", f"ไม่สามารถสแกนโฟลเดอร์ได้:\n{e}"); self.update_status("การสแกนล้มเหลว")
        finally:
            self.scan_button.config(state=tk.NORMAL)
        if not unique_subfolders: ttk.Label(self.subfolder_checkbox_frame, text="ไม่พบโฟลเดอร์ย่อยใดๆ").pack()
        else:
            max_cols = 5; row, col = 0, 0
            for name in sorted(list(unique_subfolders)):
                var = tk.BooleanVar(value=True)
                cb = ttk.Checkbutton(self.subfolder_checkbox_frame, text=name, variable=var)
                cb.grid(row=row, column=col, sticky=tk.W, padx=5, pady=2)
                self.subfolder_vars[name] = var; col += 1
                if col >= max_cols: col = 0; row += 1
        self.update_status(f"สแกนเสร็จสิ้น พบ {len(unique_subfolders)} โฟลเดอร์ย่อย")

    def toggle_ui_elements(self):
        task_id = self.selected_task.get()
        for i, frame in [(4, self.task_4_options_frame), (5, self.task_5_options_frame), (6, self.task_6_scope_frame)]:
            state = tk.NORMAL if task_id == i else tk.DISABLED
            for child in frame.winfo_children():
                try:
                    child.configure(state=state)
                    if isinstance(child, ttk.Frame):
                         for grandchild in child.winfo_children(): grandchild.configure(state=state)
                except tk.TclError: pass
        
        enable_subfolder_ui = (task_id in [1, 2, 3, 5, 7, 8]) or \
                              (task_id == 4 and self.search_scope.get() == 'selected') or \
                              (task_id == 6 and self.delete_empty_scope.get() == 'selected')
        scan_btn_state = tk.NORMAL if enable_subfolder_ui and self.base_path.get() else tk.DISABLED
        checkbox_state = tk.NORMAL if enable_subfolder_ui else tk.DISABLED
        self.scan_button.config(state=scan_btn_state)
        for child in self.subfolder_checkbox_frame.winfo_children(): child.configure(state=checkbox_state)

    def start_task(self):
        base_path, output_file, task_id = self.base_path.get(), self.output_file.get(), self.selected_task.get()
        if not all([base_path, output_file]): messagebox.showerror("ข้อมูลไม่ครบถ้วน", "กรุณาเลือกโฟลเดอร์เป้าหมายและตำแหน่งบันทึกไฟล์ Output ก่อน"); return
        if not os.path.isdir(base_path): messagebox.showerror("ผิดพลาด", "โฟลเดอร์เป้าหมายที่เลือกไม่มีอยู่จริง"); return

        if task_id in [5, 6, 7]:
            if not messagebox.askyesno("ยืนยันการทำงาน", "คุณแน่ใจหรือไม่ว่าจะย้ายไฟล์/โฟลเดอร์ตามที่เลือกไปที่ถังขยะ?\n\n(คุณยังสามารถกู้คืนไฟล์ได้จากถังขยะ)"):
                self.update_status("การทำงานถูกยกเลิก"); return
        if task_id == 8:
            if not messagebox.askyesno("ยืนยันการทำงาน", "คุณแน่ใจหรือไม่ว่าจะอัปเดตไฟล์ Excel?\n\nการทำงานนี้จะเขียนทับข้อมูลในคอลัมน์สถานะและโน้ต"):
                self.update_status("การทำงานถูกยกเลิก"); return

        args, selected_subfolders = [base_path, output_file, self.update_status], []
        
        # Determine if selected_subfolders are needed for the current task
        needs_subfolders = (task_id in [1, 2, 3, 5, 7, 8]) or \
                           (task_id == 4 and self.search_scope.get() == 'selected') or \
                           (task_id == 6 and self.delete_empty_scope.get() == 'selected')

        if needs_subfolders:
            selected_subfolders = [name for name, var in self.subfolder_vars.items() if var.get()]
            if not selected_subfolders:
                if task_id == 8:
                    # For task 8, if nothing is selected, use a default list for notes.
                    selected_subfolders = ['2D', '3D', 'Datasheet']
                else:
                    messagebox.showwarning("เลือกตัวเลือก", "กรุณาสแกนและเลือกโฟลเดอร์ย่อยอย่างน้อย 1 รายการ")
                    return
        
        task_map = {1: run_task_list_contents, 2: run_task_find_empty_subfolders, 3: run_task_summarize_types, 4: run_task_find_specific_files, 5: run_task_delete_files_by_ext, 6: run_task_delete_empty_subfolders, 7: run_task_delete_empty_main_folders, 8: run_task_update_excel}
        
        if task_id in [1, 2, 3, 7]: args.append(selected_subfolders)
        elif task_id == 4:
            ext_list = [ext.strip().lower() for ext in self.extensions_to_find.get().split(',') if ext.strip()]
            if not ext_list or not all(e.startswith('.') for e in ext_list): messagebox.showerror("รูปแบบผิดพลาด", "นามสกุลไฟล์สำหรับ Task 4 ต้องขึ้นต้นด้วยจุด (.)"); return
            args.extend([ext_list, self.search_scope.get(), selected_subfolders])
        elif task_id == 5:
            ext_list = [ext.strip().lower() for ext in self.extensions_to_delete.get().split(',') if ext.strip()]
            if not ext_list or not all(e.startswith('.') for e in ext_list): messagebox.showerror("รูปแบบผิดพลาด", "นามสกุลไฟล์สำหรับ Task 5 ต้องขึ้นต้นด้วยจุด (.)"); return
            args.extend([selected_subfolders, ext_list])
        elif task_id == 6:
            args.extend([self.delete_empty_scope.get(), selected_subfolders])
        elif task_id == 8:
            excel_input = self.excel_input_path.get()
            id_col_val, status_col_val, notes_col_val = self.id_col.get(), self.status_col.get(), self.notes_col.get()
            if not excel_input: messagebox.showerror("ข้อมูลไม่ครบถ้วน", "กรุณาเลือกไฟล์ Excel ต้นฉบับสำหรับ Task 8"); return
            if not all([id_col_val, status_col_val, notes_col_val]): messagebox.showerror("ข้อมูลไม่ครบถ้วน", "กรุณากำหนดคอลัมน์ทั้งหมดสำหรับ Task 8"); return
            args.extend([excel_input, selected_subfolders, id_col_val, status_col_val, notes_col_val])

        target_function = partial(task_map[task_id], *args)
        
        self.run_button.config(state=tk.DISABLED)
        self.update_status("กำลังเริ่มต้น...")
        threading.Thread(target=self.run_task_in_thread, args=(target_function, output_file)).start()

    def run_task_in_thread(self, target_function, output_file):
        try:
            target_function()
            self.update_status(f"เสร็จสิ้น! Log/ไฟล์ ถูกบันทึกที่ {os.path.basename(output_file)}")
            messagebox.showinfo("เสร็จสิ้น", f"การทำงานเสร็จสมบูรณ์\nไฟล์ Output ถูกบันทึกไว้ที่:\n{output_file}")
        except Exception as e:
            self.update_status(f"เกิดข้อผิดพลาด: {e}")
            messagebox.showerror("เกิดข้อผิดพลาด", f"เกิดข้อผิดพลาดระหว่างการทำงาน:\n{e}")
        finally:
            self.run_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()