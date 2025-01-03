import tkinter as tk
from tkinter import filedialog, messagebox
from app import main  # Assuming your script is in a file named app.py
import openpyxl

def run_script():
    search_input_text = search_input.get().strip()
    file_path = file_path_entry.get()
    sheet_name = sheet_name_var.get()
    lesson_number = lesson_number_var.get()
    class_performance_text = class_performance_input.get("1.0", tk.END).strip()
    homework_result_text = homework_result_input.get("1.0", tk.END).strip()
    deadline_text = deadline_input.get("1.0", tk.END).strip()
    next_requirement_text = next_requirement_input.get("1.0", tk.END).strip()

    # Call the main function from your script with the collected inputs
    try:
        main(search_input_text, file_path, sheet_name, lesson_number, class_performance_text, homework_result_text, deadline_text, next_requirement_text)
        messagebox.showinfo("Thành công", "Chạy script thành công!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {e}")

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, filename)
    update_sheet_names(filename)

def update_sheet_names(filename):
    workbook = openpyxl.load_workbook(filename, read_only=True)
    sheet_names = workbook.sheetnames
    sheet_name_var.set(sheet_names[0])  # Set the first sheet as default
    sheet_name_menu['menu'].delete(0, 'end')
    for name in sheet_names:
        sheet_name_menu['menu'].add_command(label=name, command=tk._setit(sheet_name_var, name))

# Create the main window
root = tk.Tk()
root.title("Giao diện Script")

# Create and place the widgets
tk.Label(root, text="Tìm kiếm:").grid(row=0, column=0, sticky=tk.W)
search_input = tk.Entry(root, width=50)
search_input.grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="Đường dẫn file:").grid(row=1, column=0, sticky=tk.W)
file_path_entry = tk.Entry(root, width=50)
file_path_entry.grid(row=1, column=1, padx=5, pady=5)
tk.Button(root, text="Chọn file", command=browse_file).grid(row=1, column=2, padx=5, pady=5)

tk.Label(root, text="Tên sheet:").grid(row=2, column=0, sticky=tk.W)
sheet_name_var = tk.StringVar()
sheet_name_menu = tk.OptionMenu(root, sheet_name_var, "")
sheet_name_menu.grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Số bài học:").grid(row=3, column=0, sticky=tk.W)
lesson_number_var = tk.StringVar(value="1")  # Default value
lesson_number_menu = tk.OptionMenu(root, lesson_number_var, *range(1, 15))
lesson_number_menu.grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="Hiệu suất lớp học:").grid(row=4, column=0, sticky=tk.W)
class_performance_input = tk.Text(root, height=2, width=50)
class_performance_input.grid(row=4, column=1, padx=5, pady=5)

tk.Label(root, text="Kết quả bài tập về nhà:").grid(row=5, column=0, sticky=tk.W)
homework_result_input = tk.Text(root, height=2, width=50)
homework_result_input.grid(row=5, column=1, padx=5, pady=5)

tk.Label(root, text="Hạn chót:").grid(row=6, column=0, sticky=tk.W)
deadline_input = tk.Text(root, height=2, width=50)
deadline_input.grid(row=6, column=1, padx=5, pady=5)

tk.Label(root, text="Yêu cầu tiếp theo:").grid(row=7, column=0, sticky=tk.W)
next_requirement_input = tk.Text(root, height=2, width=50)
next_requirement_input.grid(row=7, column=1, padx=5, pady=5)

tk.Button(root, text="Chạy Script", command=run_script).grid(row=8, column=1, pady=10)

# Start the main loop
root.mainloop()