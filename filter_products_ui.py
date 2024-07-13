import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def load_excel(file_path):
    df = pd.read_excel(file_path)
    return df

def find_column(df, column_name):
    if column_name in df.columns:
        return df
    else:
        raise ValueError(f"Column {column_name} not found in the Excel file")

def filter_by_category(df, category_column, category):
    filtered_df = df[df[category_column] == category]
    return filtered_df

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    file_path.set(filename)
    load_columns(filename)

def load_columns(file_path):
    df = load_excel(file_path)
    column_combobox['values'] = df.columns.tolist()

def filter_data():
    try:
        file_path_value = file_path.get()
        category_column_value = category_column.get()
        category_value = category.get()
        
        df = load_excel(file_path_value)
        df = find_column(df, 'CODIGO')
        filtered_df = filter_by_category(df, category_column_value, category_value)
        
        result_text.set(filtered_df.to_string(index=False))
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
category = tk.StringVar()
result_text = tk.StringVar()

# Set the common width
common_width = 60

# Create and place the widgets
tk.Label(root, text="Excel File:", bg="#f0f0f0", anchor="w").grid(row=0, column=0, padx=10, pady=10, sticky="ew")

file_frame = tk.Frame(root)
file_frame.grid(row=0, column=1, padx=10, pady=10, columnspan=2, sticky="ew")

file_entry = tk.Entry(file_frame, textvariable=file_path, width=common_width + 10, justify='left', font=('Arial', 10))
file_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

browse_button = tk.Button(file_frame, text="Browse", command=browse_file)
browse_button.pack(side=tk.RIGHT)

tk.Label(root, text="Column:", bg="#f0f0f0", anchor="w").grid(row=1, column=0, padx=10, pady=10, sticky="ew")
column_combobox = ttk.Combobox(root, textvariable=category_column, width=common_width - 2, justify='left')
column_combobox.grid(row=1, column=1, padx=10, pady=10, columnspan=2, sticky="ew")

tk.Label(root, text="Category:", bg="#f0f0f0", anchor="w").grid(row=2, column=0, padx=10, pady=10, sticky="ew")
category_entry = tk.Entry(root, textvariable=category, width=common_width, justify='left', font=('Arial', 10))
category_entry.grid(row=2, column=1, padx=10, pady=10, columnspan=2, sticky="ew")

# Result frame
tk.Label(root, text="Filtered Results:", bg="#f0f0f0", anchor="w").grid(row=3, column=0, padx=10, pady=10, sticky="ew")
result_frame = tk.Frame(root, bg="white", relief="sunken", borderwidth=2)
result_frame.grid(row=3, column=1, padx=10, pady=10, columnspan=3, sticky="nsew")

result_message = tk.Message(result_frame, textvariable=result_text, width=600, bg="white", anchor="nw")
result_message.pack(fill=tk.BOTH, expand=True)

# Buttons
button_frame = tk.Frame(root, bg="#f0f0f0")
button_frame.grid(row=4, column=1, columnspan=3, padx=10, pady=10, sticky="e")

tk.Button(button_frame, text="Filter", command=filter_data, bg="#2196F3", fg="white").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Export to TXT", command=export_to_txt, bg="#FF9800", fg="white").pack(side=tk.LEFT, padx=5)

# Configure grid weights for resizing
root.grid_columnconfigure(1, weight=1)
root.grid_columnconfigure(2, weight=1)
root.grid_rowconfigure(3, weight=1)

# Run the application
root.mainloop()
