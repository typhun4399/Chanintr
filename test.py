import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
import threading
import time
import pandas as pd

steps = []  # เก็บทุก Step
last_selected_index = None  # เก็บ index ล่าสุดที่เลือกใน Listbox

# ---------------- GUI ----------------
root = tk.Tk()
root.title("Selenium Step Recorder with Excel Row Input")
root.geometry("900x600")

# ---------------- Step List ----------------
def update_steps_list():
    step_listbox.delete(0, tk.END)
    for i, step in enumerate(steps, 1):
        if step["action"] == "Open URL":
            step_listbox.insert(tk.END, f"Step {i}: Open URL -> {step.get('url','')}")
        elif step["action"] == "Click":
            step_listbox.insert(tk.END, f"Step {i}: Click -> {step.get('xpath','')} (Text: {step.get('text','')})")
        elif step["action"] == "Input Text":
            if step.get("type")=="text":
                step_listbox.insert(tk.END, f"Step {i}: Input Text -> '{step.get('text','')}' at {step.get('xpath','')}")
            else:
                step_listbox.insert(tk.END, f"Step {i}: Input Text from Excel -> {step.get('excel_path','')} Column: {step.get('column','')} at {step.get('xpath','')}")
        elif step["action"] == "Loop":
            step_listbox.insert(tk.END, f"Step {i}: Loop Start -> Repeat continuously after Step {step['start_index']+1}")

def toggle_selection(event):
    global last_selected_index
    clicked_index = step_listbox.nearest(event.y)
    if last_selected_index == clicked_index:
        step_listbox.selection_clear(0, tk.END)
        last_selected_index = None
    else:
        step_listbox.selection_clear(0, tk.END)
        step_listbox.selection_set(clicked_index)
        last_selected_index = clicked_index

# ---------------- Add Step ----------------
def get_xpath_script():
    return """
    function getElementXPath(element) {
        if (!element) return null;
        if (element.id !== '') return '//*[@id="' + element.id + '"]';
        if (element === document.body) return '/html/body';
        var ix = 0;
        var siblings = element.parentNode ? element.parentNode.childNodes : [];
        for (var i=0; i<siblings.length; i++) {
            var sibling = siblings[i];
            if (sibling === element) {
                return getElementXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix+1) + ']';
            }
            if (sibling.nodeType === 1 && sibling.tagName === element.tagName) {
                ix++;
            }
        }
        return null;
    }
    window.xpath_result = null;
    ["click","mousedown","mouseup"].forEach(evt=>{
        document.addEventListener(evt,function(e){
            var path = e.composedPath ? e.composedPath()[0] : e.target;
            var target = path.closest ? path.closest('*') : path;
            if(!target) return;
            var xp = getElementXPath(target);
            var txt = target.innerText || target.getAttribute("aria-label") || target.getAttribute("alt") || "";
            window.xpath_result={xpath:xp,text:txt};
        },{capture:true});
    });
    """

def add_step_action(action_combo):
    action = action_combo.get()
    selected_index = step_listbox.curselection()
    insert_index = selected_index[0] if selected_index else len(steps)

    if action == "Open URL":
        url = simpledialog.askstring("Input URL", "กรอก URL:", parent=root)
        if not url: return
        steps.insert(insert_index, {"action": "Open URL", "url": url})
        update_steps_list()
        return

    if action == "Input Text":
        last_click = None
        for step in reversed(steps[:insert_index]):
            if step["action"] == "Click":
                last_click = step
                break
        if not last_click:
            messagebox.showwarning("Warning", "ต้องมี Step Click เลือก element ก่อน Input Text", parent=root)
            return

        # ---------------- Popup เลือก Text หรือ Excel แบบซ้ายขวา ----------------
        type_popup = tk.Toplevel(root)
        type_popup.title("เลือกประเภท Input Text")
        tk.Label(type_popup, text="เลือกประเภท Input Text:").pack(pady=10)

        button_frame = tk.Frame(type_popup)
        button_frame.pack(pady=5)

        def choose_text():
            type_popup.destroy()
            input_text_value = simpledialog.askstring("Input Text", "กรอกข้อความที่จะใส่:", parent=root)
            if input_text_value is None: return
            steps.insert(insert_index, {
                "action": "Input Text",
                "xpath": last_click["xpath"],
                "text": input_text_value,
                "url": last_click.get("url", ""),
                "type": "text"
            })
            update_steps_list()

        def choose_excel():
            type_popup.destroy()
            excel_file = filedialog.askopenfilename(title="เลือกไฟล์ Excel", filetypes=[("Excel files","*.xlsx *.xls")])
            if not excel_file: return
            df = pd.read_excel(excel_file)
            columns = list(df.columns)
            if not columns:
                messagebox.showwarning("Warning","Excel ไม่มีคอลัมน์", parent=root)
                return

            column_win = tk.Toplevel(root)
            column_win.title("เลือก Column")
            tk.Label(column_win,text="เลือก Column:").pack(pady=5)
            column_combo = ttk.Combobox(column_win, values=columns, state="readonly")
            column_combo.pack(pady=5)
            column_combo.current(0)
            def select_column():
                col = column_combo.get()
                # ตอน Add Step ให้ Excel input แค่ row แรก
                steps.insert(insert_index,{
                    "action": "Input Text",
                    "xpath": last_click["xpath"],
                    "url": last_click.get("url", ""),
                    "type": "excel",
                    "excel_path": excel_file,
                    "column": col,
                    "current_row": 0  # row แรก
                })
                update_steps_list()
                column_win.destroy()
            tk.Button(column_win,text="Select", command=select_column).pack(pady=5)
            column_win.grab_set()

        tk.Button(button_frame, text="Text ปกติ", width=15, command=choose_text).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="Excel", width=15, command=choose_excel).pack(side=tk.RIGHT, padx=10)
        type_popup.grab_set()
        return

    elif action == "Loop":
        if insert_index == 0:
            messagebox.showwarning("Warning", "ต้องมี Step ก่อนหน้าเพื่อ Loop", parent=root)
            return
        # เพิ่ม Step Loop แต่ไม่ทำอะไรตอน Add Step
        steps.insert(insert_index, {"action": "Loop", "start_index": insert_index})
        update_steps_list()
        return

    else:  # Click
        url = None
        for step in reversed(steps[:insert_index]):
            if "url" in step:
                url = step["url"]
                break
        if not url:
            messagebox.showwarning("Warning", "กรุณาเพิ่ม Step Open URL ก่อน", parent=root)
            return

        def selenium_thread():
            driver = webdriver.Chrome()
            driver.maximize_window()
            driver.get(url)
            time.sleep(1)
            for step in steps[:insert_index]:
                try:
                    if step["action"] == "Click":
                        driver.find_element(By.XPATH, step["xpath"]).click()
                        time.sleep(1)
                    elif step["action"] == "Input Text" and step.get("type")=="text":
                        el = driver.find_element(By.XPATH, step["xpath"])
                        el.clear()
                        el.send_keys(step["text"])
                        time.sleep(1)
                except: pass
            driver.execute_script(get_xpath_script())
            messagebox.showinfo("Info","คลิก element สำหรับ Step ใหม่ (Click)", parent=root)
            current_xpath = None
            while True:
                try:
                    result = driver.execute_script("return window.xpath_result;")
                    if result and result["xpath"] != current_xpath:
                        current_xpath = result["xpath"]
                        steps.insert(insert_index, {"url": driver.current_url, "action": "Click", "xpath": current_xpath, "text": result["text"]})
                        update_steps_list()
                        driver.quit()
                        break
                except: pass
                time.sleep(0.2)

        threading.Thread(target=selenium_thread, daemon=True).start()

# ---------------- Delete Step ----------------
def delete_step():
    selected = step_listbox.curselection()
    if not selected:
        messagebox.showwarning("Warning", "กรุณาเลือก Step ที่ต้องการลบ", parent=root)
        return
    del steps[selected[0]]
    update_steps_list()

# ---------------- Run Script ----------------
def run_script():
    if not steps:
        messagebox.showwarning("Warning", "ไม่มี Step ให้รัน", parent=root)
        return

    def selenium_run_thread():
        driver = webdriver.Chrome()
        driver.maximize_window()
        i = 0
        while i < len(steps):
            step = steps[i]
            try:
                if step["action"] == "Open URL":
                    driver.get(step["url"])
                    time.sleep(2)
                elif step["action"] == "Click":
                    driver.find_element(By.XPATH, step["xpath"]).click()
                    time.sleep(1)
                elif step["action"] == "Input Text":
                    el = driver.find_element(By.XPATH, step["xpath"])
                    el.clear()
                    if step.get("type")=="text":
                        el.send_keys(step["text"])
                        time.sleep(1)
                    else:  # Excel row by row
                        df = pd.read_excel(step["excel_path"])
                        if "current_row" not in step:
                            step["current_row"] = 0
                        if step["current_row"] < len(df):
                            val = df[step["column"]].iloc[step["current_row"]]
                            el.send_keys(str(val))
                            step["current_row"] += 1
                            time.sleep(0.5)
                elif step["action"] == "Loop":
                    # ทำงานเฉพาะตอน Run Script
                    while True:
                        for j in range(step["start_index"], len(steps)):
                            inner = steps[j]
                            try:
                                if inner["action"]=="Click":
                                    driver.find_element(By.XPATH, inner["xpath"]).click()
                                elif inner["action"]=="Input Text":
                                    el = driver.find_element(By.XPATH, inner["xpath"])
                                    el.clear()
                                    if inner.get("type")=="text":
                                        el.send_keys(inner["text"])
                                    else:
                                        df = pd.read_excel(inner["excel_path"])
                                        if "current_row" not in inner:
                                            inner["current_row"] = 0
                                        if inner["current_row"] < len(df):
                                            val = df[inner["column"]].iloc[inner["current_row"]]
                                            el.send_keys(str(val))
                                            inner["current_row"] += 1
                                time.sleep(1)
                            except: pass
            except Exception as e:
                print("Run error:", e)
            i += 1
        messagebox.showinfo("Info", "Run script เสร็จเรียบร้อย (browser ยังเปิดอยู่)", parent=root)

    threading.Thread(target=selenium_run_thread, daemon=True).start()

# ---------------- GUI Layout ----------------
top_frame = tk.Frame(root)
top_frame.pack(pady=10)
tk.Label(top_frame, text="Add Step:").pack(side=tk.LEFT, padx=5)
action_combo = ttk.Combobox(top_frame, values=["Open URL","Click","Input Text","Loop"], state="readonly", width=15)
action_combo.pack(side=tk.LEFT, padx=5)
action_combo.current(0)
tk.Button(top_frame, text="Add Step", command=lambda: add_step_action(action_combo)).pack(side=tk.LEFT, padx=5)
tk.Button(top_frame, text="Exit", command=root.destroy).pack(side=tk.LEFT, padx=10)

step_listbox = tk.Listbox(root, width=120)
step_listbox.pack(pady=10)
step_listbox.bind("<Button-1>", toggle_selection)

bottom_frame = tk.Frame(root)
bottom_frame.pack(pady=10)
tk.Button(bottom_frame, text="Delete Step", command=delete_step).pack(side=tk.LEFT, padx=10)
tk.Button(bottom_frame, text="Run Script", command=run_script).pack(side=tk.LEFT, padx=10)

root.mainloop()
