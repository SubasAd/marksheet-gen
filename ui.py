import tkinter as tk
from tkinter import filedialog, messagebox
from excel_reader import read_students,read_school_info
from data_mapper import map_marksheet_data_by_section
from doc_builder.doc_builder_terminal import create_marksheet_docx
from doc_builder.doc_builder_final import create_marksheet_docx as final_marksheet_generator
from config import START_ROW

def browse_file(entry_widget):
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, filename)

def browse_logo(entry_widget):
    filename = filedialog.askopenfilename(filetypes=[("Image Files", ["*.png","*.jpg","*.jpeg"])])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, filename)

def generate_marksheets():
    try:
    
        school_details = read_school_info(ledger_entry.get())
        if class_entry.get():
         school_details["grade"] = class_entry.get()
        if exam_bs_entry.get():
         school_details["exam_year_bs"] = exam_bs_entry.get()
        if exam_ad_entry.get():
         school_details["exam_year_ad"] = exam_ad_entry.get()
        sheet_name = "Class" + str(school_details["grade"])
        students = read_students(sheet_name, ledger_entry.get(), start_row=START_ROW)
        filtered = [s for s in students if s.get("2 - Roll No.") == int(roll_entry.get())]

        if not filtered:
            messagebox.showerror("Error", "No student found with this roll number!")
            return

        for flat_data in filtered:
            structured = map_marksheet_data_by_section(flat_data)
            create_marksheet_docx(structured, "first_term", school_details, f"First_Term_Marksheet{school_details["grade"]}-{roll_entry.get()}.docx")
            create_marksheet_docx(structured, "second_term", school_details, f"Second_Term_Marksheet{school_details["grade"]}-{roll_entry.get()}.docx")
            final_marksheet_generator(structured, school_details, f"Annual_Marksheet{school_details["grade"]}-{roll_entry.get()}.docx")
        messagebox.showinfo("Success", "Marksheets generated successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Marksheet Generator")

tk.Label(root, text="Ledger File:").grid(row=0, column=0, sticky="e")
ledger_entry = tk.Entry(root, width=50)
ledger_entry.grid(row=0, column=1)
tk.Button(root, text="Browse", command=lambda: browse_file(ledger_entry)).grid(row=0, column=2)

tk.Label(root, text="Roll Number:").grid(row=1, column=0, sticky="e")
roll_entry = tk.Entry(root)
roll_entry.grid(row=1, column=1)

tk.Label(root, text="Class:").grid(row=2, column=0, sticky="e")
class_entry = tk.Entry(root)
class_entry.grid(row=2, column=1)

tk.Label(root, text="Exam Year (BS):").grid(row=6, column=0, sticky="e")
exam_bs_entry = tk.Entry(root)
exam_bs_entry.grid(row=6, column=1)

tk.Label(root, text="Exam Year (AD):").grid(row=7, column=0, sticky="e")
exam_ad_entry = tk.Entry(root)
exam_ad_entry.grid(row=7, column=1)


tk.Button(root, text="Generate Marksheets", command=generate_marksheets).grid(row=8, column=0, columnspan=3, pady=10)

root.mainloop()
