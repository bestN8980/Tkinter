import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import asksaveasfilename
from tkcalendar import DateEntry
import pandas as pd
import csv
from datetime import datetime

# Hàm lưu thông tin vào file CSV
def save_to_csv():
    data = {
        "Mã": entry_id.get(),
        "Tên": entry_name.get(),
        "Đơn vị": combobox.get(),
        "Chức danh": entry_position.get(),
        "Ngày sinh": entry_birth.get_date().strftime('%d/%m/%Y'),
        "Giới tính": gender_var.get(),
        "Số CMND": entry_id_number.get(),
        "Ngày cấp": entry_issue_date.get_date().strftime('%d/%m/%Y'),
        "Nơi cấp": entry_issue_place.get()
    }

    if all(data.values()):
        with open("employee_data.csv", mode="a", newline='', encoding="utf-8") as file:
            writer = csv.DictWriter(file, fieldnames=data.keys())
            if file.tell() == 0:
                writer.writeheader()
            writer.writerow(data)
        messagebox.showinfo("Thành công", "Dữ liệu đã được lưu thành công.")
    else:
        messagebox.showerror("Lỗi", "Vui lòng điền đầy đủ thông tin.")

# Hàm hiển thị nhân viên có sinh nhật hôm nay
def show_today_birthdays():
    try:
        today = datetime.today().strftime('%d/%m/%Y')
        with open("employee_data.csv", mode="r", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            birthdays = [row for row in reader if row['Ngày sinh'] == today]

        if birthdays:
            result = "\n".join([f"Mã: {row['Mã']}, Tên: {row['Tên']}, Ngày sinh: {row['Ngày sinh']}" for row in birthdays])
            messagebox.showinfo("Danh sách sinh nhật hôm nay", result)
        else:
            messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay.")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu. Vui lòng nhập thông tin trước.")

# Hàm xuất toàn bộ danh sách ra Excel
def export_to_excel():
    try:
        df = pd.read_csv("employee_data.csv")
        df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format='%d/%m/%Y')
        df = df.sort_values(by=['Ngày sinh'], ascending=False)

        filepath = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if filepath:
            df.to_excel(filepath, index=False)
            messagebox.showinfo("Thành công", f"Dữ liệu đã được xuất ra file {filepath}")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu để xuất.")

# Giao diện chính
root = tk.Tk()
root.title("Thông tin nhân viên")
root.geometry('900x300')

# Các biến lưu trữ thông tin
entry_id = tk.StringVar()
entry_name = tk.StringVar()
entry_department = tk.StringVar()
entry_position = tk.StringVar()
gender_var = tk.StringVar(value="Nam")
entry_id_number = tk.StringVar()
entry_issue_place = tk.StringVar()

# Tạo các widget
frame = ttk.Frame(root)
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Các trường nhập liệu
labels = [("Mã", " *"), ("Tên", " *"), ("Đơn vị", " *"), ("Chức danh", ""), ("Ngày sinh", ""), ("Giới tính", ""), ("Số CMND", ""), ("Ngày cấp", ""), ("Nơi cấp", "")]
entries = [entry_id, entry_name, entry_department, entry_position, None, None, entry_id_number, None, entry_issue_place]

#Dòng đầu
tk.Label(frame, text="THÔNG TIN NHÂN VIÊN", font=("Arial", 25)).grid(row=0, column=0,columnspan=2, padx=10, sticky=tk.W)
checkbox1 = ttk.Checkbutton(frame, text="Là khách hàng")
checkbox2 = ttk.Checkbutton(frame, text="Là nhà cung cấp")

checkbox1.grid(row=0, column=2, padx=10)
checkbox2.grid(row=0, column=3, padx=10)

# Mã
tk.Label(frame, text=labels[0][0], font=("Arial", 10)).grid(row=1, column=0, padx=10, sticky=tk.W)
tk.Label(frame, text=labels[0][1], font=("Arial", 10), fg="red").grid(row=1, column=0, padx=10)
ttk.Entry(frame, textvariable=entries[0], width=30).grid(row=2, column=0, padx=10, sticky=tk.W)

# Tên
tk.Label(frame, text=labels[1][0], font=("Arial", 10)).grid(row=1, column=1, padx=0, sticky=tk.W)
tk.Label(frame, text=labels[1][1], font=("Arial", 10), fg="red").grid(row=1, column=1, padx=10)
ttk.Entry(frame, textvariable=entries[1], width=30).grid(row=2, column=1, padx=0, sticky=tk.W)

# Ngày sinh
tk.Label(frame, text=labels[4][0], font=("Arial", 10)).grid(row=1, column=2, padx=20, sticky=tk.W)
entry_birth = DateEntry(frame, date_pattern='dd/MM/yyyy', font=("Arial", 10))
entry_birth.grid(row=2, column=2, padx=20)

# Giới tính
tk.Label(frame, text=labels[5][0], font=("Arial", 10)).grid(row=1, column=3, sticky=tk.W)
gender_frame = ttk.Frame(frame)
gender_frame.grid(row=2, column=3, sticky=(tk.W, tk.E), pady=5)
ttk.Radiobutton(gender_frame, text="Nam", variable=gender_var, value="Nam").pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(gender_frame, text="Nữ", variable=gender_var, value="Nữ").pack(side=tk.LEFT, padx=5)

#Đơn vị
def show_selection(event): 
    selected = combobox.get() 
    event.config(text=f"{selected}")

tk.Label(frame, text=labels[2][0], font=("Arial", 10)).grid(row=3, column=0, sticky=tk.W, padx=10)
options = ["Phân xưởng 1", "Phân xưởng 2", "Phân xưởng 3", "Phân xưởng 4"]
combobox = ttk.Combobox(frame, values=options, state="readonly", width=61) 
combobox.grid(row=4, column=0,columnspan=2, padx=10, sticky=tk.W) 
combobox.bind("<<ComboboxSelected>>", show_selection)

# CMND
tk.Label(frame, text=labels[6][0], font=("Arial", 10)).grid(row=3, column=2, columnspan=2, sticky=tk.W, padx=20)
ttk.Entry(frame, textvariable=entries[6], width=30).grid(row=4, column=2, columnspan=2, sticky=tk.W, padx=20)

# Ngày cấp
tk.Label(frame, text=labels[7][0], font=("Arial", 10)).grid(row=3, column=3,columnspan=2, padx=10, sticky=tk.W)
entry_issue_date = DateEntry(frame, date_pattern='dd/MM/yyyy', font=("Arial", 10))
entry_issue_date.grid(row=4, column=3, padx=10)

# Nút chức năng
button_frame = ttk.Frame(root, padding=10)
button_frame.grid(row=8, column=0, sticky=(tk.W, tk.E))

#Chức danh
tk.Label(frame, text=labels[3][0], font=("Arial", 10)).grid(row=5, column=0, sticky=tk.W, padx=10)
ttk.Entry(frame, textvariable=entries[3], width=64).grid(row=6, column=0, columnspan=2, sticky=tk.W, padx=10)

#Noi cấp
tk.Label(frame, text=labels[8][0], font=("Arial", 10)).grid(row=5, column=2, sticky=tk.W, padx=20)
ttk.Entry(frame, textvariable=entries[8], width=64).grid(row=6, column=2, columnspan=2, sticky=tk.W, padx=20)

ttk.Button(button_frame, text="Lưu thông tin", command=save_to_csv).grid(row=0, column=0, padx=5)
ttk.Button(button_frame, text="Sinh nhật hôm nay", command=show_today_birthdays).grid(row=0, column=1, padx=5)
ttk.Button(button_frame, text="Xuất danh sách", command=export_to_excel).grid(row=0, column=2, padx=5)

root.mainloop()
