import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def load_excel(file_path):
    df = pd.read_excel(file_path)
    df = df.loc[:, df.columns != 'Id']  # Exclude the 'id' column
    return df

def load_columns(file_path):
    df = load_excel(file_path)
    column_combobox['values'] = df.columns.tolist()
    if not df.empty:
        category_column.set(df.columns[0])  # Automatically select the first column
        update_categories(None)  # Update categories for the first column

def update_categories(event):
    try:
        file_path_value = file_path.get()
        df = load_excel(file_path_value)
        selected_column = category_column.get()
        unique_categories = df[selected_column].unique()
        category_combobox['values'] = unique_categories.tolist()
        if unique_categories.size > 0:
            category_combobox.set(unique_categories[0])  # Automatically select the first category
            filter_data()  # Automatically filter data
    except Exception as e:
        messagebox.showerror("Error", str(e))

def filter_by_category(df, category_column, category):
    filtered_df = df[df[category_column] == category]
    return filtered_df

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    file_path.set(filename)
    load_columns(filename)

def filter_data():
    try:
        file_path_value = file_path.get()
        category_column_value = category_column.get()
        category_value = category_combobox.get()
        
        df = load_excel(file_path_value)
        filtered_df = filter_by_category(df, category_column_value, category_value)

        # Clear the Treeview
        for item in tree.get_children():
            tree.delete(item)

        # Insert new rows into the Treeview
        if not filtered_df.empty:
            tree["columns"] = filtered_df.columns.tolist()
            for col in filtered_df.columns:
                tree.heading(col, text=col)
                tree.column(col, anchor="center", width=120)

            for index, row in filtered_df.iterrows():
                tree.insert("", "end", values=row.tolist())
        else:
            messagebox.showinfo("Info", "No results found for the selected category.")

        global filtered_df_global
        filtered_df_global = filtered_df
        
    except Exception as e:
        messagebox.showerror("Error", str(e))

def export_to_txt():
    try:
        export_file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if export_file_path:
            with open(export_file_path, 'w') as f:
                f.write(filtered_df_global.to_string(index=False))
            messagebox.showinfo("Success", f"Results exported to {export_file_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Create the main window
root = tk.Tk()
root.title("Excel Category Filter")
root.geometry("800x500")
root.configure(bg="#f0f0f0")

# Create StringVar variables
file_path = tk.StringVar()
category_column = tk.StringVar()

# Set the common width
common_width = 75

# Create and place the widgets
tk.Label(root, text="Excel File:", bg="#f0f0f0", anchor="w").grid(row=0, column=0, padx=10, pady=10, sticky="ew")

file_frame = tk.Frame(root)
file_frame.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

file_entry = tk.Entry(file_frame, textvariable=file_path, width=common_width + 10, justify='left', font=('Arial', 10))
file_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)

browse_button = tk.Button(file_frame, text="Browse", width=6, command=browse_file)
browse_button.pack(side=tk.LEFT, expand=False)

tk.Label(root, text="Category Column:", bg="#f0f0f0", anchor="w").grid(row=1, column=0, padx=10, pady=10, sticky="ew")
column_combobox = ttk.Combobox(root, textvariable=category_column, width=common_width - 2, justify='left')
column_combobox.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
column_combobox.bind("<<ComboboxSelected>>", update_categories)

tk.Label(root, text="Category:", bg="#f0f0f0", anchor="w").grid(row=2, column=0, padx=10, pady=10, sticky="ew")
category_combobox = ttk.Combobox(root, width=common_width, justify='left')
category_combobox.grid(row=2, column=1, padx=10, pady=10, sticky="ew")


# Result frame
tk.Label(root, text="Filtered Results:", bg="#f0f0f0", anchor="w").grid(row=3, column=0, padx=10, pady=10, sticky="ew")
result_frame = tk.Frame(root, bg="white", relief="sunken", borderwidth=2, width=700, height=300)
result_frame.grid(row=3, column=1, padx=10, pady=10, columnspan=3, sticky="nsew")

# Scrollbars
vsb = ttk.Scrollbar(result_frame, orient="vertical")
vsb.pack(side=tk.RIGHT, fill='y')

hsb = ttk.Scrollbar(result_frame, orient="horizontal")
hsb.pack(side=tk.BOTTOM, fill='x')

# Treeview for displaying results
tree = ttk.Treeview(result_frame, columns=[], show='headings', yscrollcommand=vsb.set, xscrollcommand=hsb.set)
tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)

# Fix width of Treeview columns
tree.column("#0", width=700, stretch=tk.NO)  # Hide the first empty column

# Set fixed width for columns
def set_column_widths():
    for col in tree["columns"]:
        tree.column(col, width=700, anchor="center")

tree.bind("<Configure>", lambda e: set_column_widths())

vsb.config(command=tree.yview)
hsb.config(command=tree.xview)

# Ensure the scroll region is updated
def update_scroll_region(event):
    tree.config(scrollregion=tree.bbox("all"))

tree.bind("<Configure>", update_scroll_region)



# Buttons
button_frame = tk.Frame(root, bg="#f0f0f0")
button_frame.grid(row=5, column=1, columnspan=3, padx=10, pady=10, sticky="e")

tk.Button(button_frame, text="Filter", command=filter_data, bg="#2196F3", fg="white").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Export to TXT", command=export_to_txt, bg="#FF9800", fg="white").pack(side=tk.LEFT, padx=5)

# Run the application
root.mainloop()
