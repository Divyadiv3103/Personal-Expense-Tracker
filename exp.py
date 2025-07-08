import pyodbc
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

def db_connect():
    return pyodbc.connect(
        r"Driver={SQL SERVER};Server=LAPTOP-NS9TJCK5;Database=ExpenseDB;Trusted_Connection=yes"
    )

def add_expense():
    date = entry_date.get()
    category = entry_category.get()
    amount = entry_amount.get()

    if not(date and category and amount):
        messagebox.showerror("Error", "Expense Date, category, amount is required")
        return
    
    try:
        amount = float(amount)
    except ValueError as v:
        messagebox.showerror("Error", f"Amount must be a number: {v}") 
        return

    try:
        connection = db_connect()
        cursor = connection.cursor()
        cursor.execute(
            "INSERT INTO TR.EXPENSES (EXPENSE_DATE, CATEGORY, EXPENSE_AMOUNT) VALUES (?, ?, ?)",
           (date, category, amount)
        )
        connection.commit()
        connection.close()
        view_expense()
        clear_entry()
        messagebox.showinfo("Success", f"Expense Added Successfully: {category} - {amount}")
    except Exception as e:
        messagebox.showerror("ERROR", e)

def view_expense():
    for row in tree.get_children():
        tree.delete(row)

    try:
        connection = db_connect()
        cursor = connection.cursor()
        cursor.execute(
            "SELECT ID, EXPENSE_DATE, CATEGORY, EXPENSE_AMOUNT FROM TR.EXPENSES ORDER BY EXPENSE_DATE DESC"
        )
        for index, row in enumerate(cursor.fetchall()):
            id, date, category, amount = row
            tree.insert('', 'end', values=(
                id,
                str(date),
                category.strip(),   
                float(amount)
            ), tags=('evenrow' if index % 2 == 0 else 'oddrow',))
        connection.close()
    except Exception as e:
        messagebox.showerror("ERROR", e)

def del_expense():
    id = entry_id.get()

    try: 
        connection = db_connect()
        cursor = connection.cursor()
        cursor.execute(
           "DELETE FROM TR.EXPENSES WHERE ID = ?", (int(id),)
        )
        if cursor.rowcount == 0:
            messagebox.showwarning("Warning", f"No records found with ID: {int(id)}")
        else:
            connection.commit()
            connection.close()
            view_expense()
            entry_id.delete(0, tk.END)
            messagebox.showinfo("Deletion Success", f"Successfully deleted expense ID: {int(id)}")
    except Exception as d:
        messagebox.showerror("DB ERROR", f"Error Deleting the expense: {d}")

def upd_expense():
    selected = tree.focus()
    if not selected:
        messagebox.showerror("Error", "Select a record to update")
        return
    
    values = tree.item(selected, 'values')
    exp_id, old_date, old_category, old_amount = values

    upd_pop = tk.Toplevel(window)
    upd_pop.geometry("250x220")
    upd_pop.title("Update Expense")

    tk.Label(upd_pop, text="Id").grid(row=0, column=0, padx=10, pady=10)
    entry_upd_id = tk.Entry(upd_pop)
    entry_upd_id.grid(row=0, column=1, padx=10, pady=10)
    entry_upd_id.insert(0, exp_id)
    entry_upd_id.config(state="readonly")

    tk.Label(upd_pop, text="Date").grid(row=1, column=0, padx=10, pady=10)
    entry_upd_date = tk.Entry(upd_pop)
    entry_upd_date.grid(row=1, column=1, padx=10, pady=10)
    entry_upd_date.insert(0, old_date)

    tk.Label(upd_pop, text="Category").grid(row=2, column=0, padx=10, pady=10)
    entry_upd_category = tk.Entry(upd_pop)
    entry_upd_category.grid(row=2, column=1, padx=10, pady=10)
    entry_upd_category.insert(0, old_category)

    tk.Label(upd_pop, text="Amount").grid(row=3, column=0, padx=10, pady=10)
    entry_upd_amount = tk.Entry(upd_pop)
    entry_upd_amount.grid(row=3, column=1, padx=10, pady=10)
    entry_upd_amount.insert(0, old_amount)

    def upd_data():
        new_date = entry_upd_date.get().strip() or old_date
        new_category = entry_upd_category.get().strip() or old_category
        new_amount = entry_upd_amount.get().strip() or old_amount

        try:
            new_amount = float(new_amount)
        except ValueError:
            messagebox.showerror("Input Error", "Amount must be a number")
            return

        try:
            connection = db_connect()
            cursor = connection.cursor()
            cursor.execute(
                "UPDATE TR.EXPENSES SET EXPENSE_DATE = ?, CATEGORY = ?, EXPENSE_AMOUNT = ? WHERE ID = ?",
                (new_date, new_category, new_amount, exp_id)
            )
            connection.commit()
            connection.close()
            upd_pop.destroy()
            view_expense() 
            messagebox.showinfo("Success", f"Expense ID: {exp_id} updated successfully")
        except Exception as e:
            messagebox.showerror("Update Error", str(e))

    tk.Button(upd_pop, text="Update Data", command=upd_data, bg="#2196F3", fg="white").grid(row=4, column=1, pady=10)

def clear_entry():
    entry_date.delete(0, tk.END)
    entry_category.delete(0, tk.END)
    entry_amount.delete(0, tk.END)


window = tk.Tk()
window.title("Personal Expense Tracker")
window.geometry("800x700")
window.configure(bg="#F0F8FF")

project_title = tk.Label(
    window,
    text="💰 Personal Expense Tracker",
    font=("Helvetica", 24, "bold"),
    fg="darkblue",
    bg="#F0F8FF"
)
project_title.pack(pady=(30, 15))

sub_frame = tk.Frame(window, bg="#F0F8FF")
sub_frame.pack(padx=20, pady=20)


tk.Label(sub_frame, text="Date (YYYY-MM-DD):", font=("Arial", 11), bg="#F0F8FF").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_date = tk.Entry(sub_frame, font=("Arial", 11))
entry_date.grid(row=0, column=1, padx=5, pady=5)

tk.Label(sub_frame, text="ID:", font=("Arial", 11), bg="#F0F8FF").grid(row=0, column=2, padx=5, pady=5, sticky="e")
entry_id = tk.Entry(sub_frame, font=("Arial", 11))
entry_id.grid(row=0, column=3, padx=5, pady=5)

tk.Label(sub_frame, text="Category:", font=("Arial", 11), bg="#F0F8FF").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entry_category = tk.Entry(sub_frame, font=("Arial", 11))
entry_category.grid(row=1, column=1, padx=5, pady=5)

tk.Label(sub_frame, text="Amount:", font=("Arial", 11), bg="#F0F8FF").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_amount = tk.Entry(sub_frame, font=("Arial", 11))
entry_amount.grid(row=2, column=1, padx=5, pady=5)


tk.Button(sub_frame, text="Add Expense", command=add_expense, bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).grid(row=3, column=1, padx=10, pady=10)
tk.Button(sub_frame, text="Delete Expense", command=del_expense, bg="#f44336", fg="white", font=("Arial", 10, "bold")).grid(row=2, column=3, padx=10, pady=10)
tk.Button(sub_frame, text="Update Expense", command=upd_expense, bg="#2196F3", fg="white", font=("Arial", 10, "bold")).grid(row=3, column=3, padx=10, pady=10)


style = ttk.Style()
style.configure("Treeview.Heading", font=("Arial", 11, "bold"), background="#D3D3D3", foreground="black")
style.configure("Treeview", font=("Arial", 10), rowheight=25, background="#F9F9F9", fieldbackground="#F9F9F9")
style.map('Treeview', background=[('selected', '#3399FF')])

tree = ttk.Treeview(window, columns=("Id", "date", "category", "amount"), show="headings")
tree.heading("Id", text="ID")
tree.heading("date", text="DATE")
tree.heading("category", text="CATEGORY")
tree.heading("amount", text="AMOUNT")
tree.pack(expand=True, fill='both', padx=20, pady=20)

tree.tag_configure('oddrow', background="#E8F6F3")
tree.tag_configure('evenrow', background="#D0ECE7")

view_expense()
window.mainloop()
