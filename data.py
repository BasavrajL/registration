import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook
from tkcalendar import DateEntry
from tkinter import ttk
def save_data():
    name = entry_name.get()
    age = entry_age.get()
    email = entry_email.get()
    terminal = entry_terminal.get()
    year = year_var.get()
    division = division_var.get()
    dob = cal.get_date()
    time = time_var.get()

    if not all([name, age, email, year, division, dob, time]):
       messagebox.showerror("Error", "Bahut Naughty Ho rahe HO!!!!.")
       return
    
    # Create or load the workbook
    try:
        wb = Workbook()
        sheet = wb.active
        sheet.append(["Name", "Age", "Email","Terminal_no","Year","Division","Date","Time"])
        sheet.append([name, age, email,terminal,year,division,dob,time])
        wb.save("data.xlsx")
        messagebox.showinfo("Data Saved", "Data saved to data.xlsx")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
def clear_fields():
    entry_name.delete(0, tk.END)
    entry_age.delete(0, tk.END)
    entry_email.delete(0, tk.END)
    entry_terminal.delete(0,tk.END)
    cal.set_date("")
    time_var.set("12:00 AM")

# Create main window
root = tk.Tk()
root.title("Data Entry Form")

college_name_label = tk.Label(root, text="Nagesh Karajagi Orchid College,Solapur", font=("Arial", 14, "bold"))
college_name_label.grid(row=0, columnspan=2, padx=10, pady=5)
lab_name_label = tk.Label(root, text="NETWORK LAB", font=("Arial", 12,"bold"))
lab_name_label.grid(row=1, columnspan=2, padx=10, pady=5)

# Create and place labels and entry fields
label_name = tk.Label(root, text="Name:")
label_name.grid(row=2, column=0, padx=10, pady=5, sticky="e")
entry_name = tk.Entry(root)
entry_name.grid(row=2, column=1, padx=10, pady=5)

label_age = tk.Label(root, text="Age:")
label_age.grid(row=3, column=0, padx=10, pady=5, sticky="e")
entry_age = tk.Entry(root)
entry_age.grid(row=3, column=1, padx=10, pady=5)

label_email = tk.Label(root, text="Email:")
label_email.grid(row=4, column=0, padx=10, pady=5, sticky="e")
entry_email = tk.Entry(root)
entry_email.grid(row=4, column=1, padx=10, pady=5)

label_terminal = tk.Label(root, text="Terminal_no:")
label_terminal.grid(row=5, column=0, padx=10, pady=5, sticky="e")
entry_terminal = tk.Entry(root)
entry_terminal.grid(row=5, column=1, padx=10, pady=5)

label_year = tk.Label(root, text="Year:")
label_year.grid(row=6, column=0, padx=10, pady=5, sticky="e")
year_var = tk.StringVar(root)
year_var.set("First Year")
year_menu = tk.OptionMenu(root, year_var, "First Year", "Second Year", "Third Year", "Fourth Year")
year_menu.grid(row=6, column=1, padx=10, pady=5)

# Drop-down menu for college division
label_division = tk.Label(root, text="Division:")
label_division.grid(row=7, column=0, padx=10, pady=5, sticky="e")
division_var = tk.StringVar(root)
division_var.set("A")
division_menu = tk.OptionMenu(root, division_var, "A", "B", "C", "D")
division_menu.grid(row=7, column=1, padx=10, pady=5)

label_dob = tk.Label(root, text="Date:")
label_dob.grid(row=8, column=0, padx=10, pady=5, sticky="e")
cal = DateEntry(root)
cal.grid(row=8, column=1, padx=10, pady=5)

label_time = tk.Label(root, text="Time:")
label_time.grid(row=9, column=0, padx=10, pady=5, sticky="e")
time_var = tk.StringVar(root)
time_var.set("12:00 AM")
time_spinbox = ttk.Combobox(root, textvariable=time_var, values=[f"{i:02d}:00 AM" for i in range(1, 13)] + [f"{i:02d}:00 PM" for i in range(1, 13)])
time_spinbox.grid(row=9, column=1, padx=10, pady=5)

# Create and place a button to save data
button_save = tk.Button(root, text="Save", command=save_data)
button_save.grid(row=10, column=1, padx=(10, 5), pady=10)

button_clear = tk.Button(root, text="Clear", command=clear_fields)
button_clear.grid(row=10, column=0, padx=(5, 10), pady=10)

# Start the Tkinter event loop
root.mainloop()
