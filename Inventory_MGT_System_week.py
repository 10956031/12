#!/usr/bin/env python
# coding: utf-8

# In[3]:


import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Initial data base
data = {
    "Product": ["Lobster", "Lobster", "Lobster", "Salmon", "Salmon", "Salmon", "Shrimp", "Shrimp", "Shrimp"],
    "Item": ["Supply", "Demand", "Inventory", "Supply", "Demand", "Inventory", "Supply", "Demand", "Inventory"],
    "Week 0": [25, 10, 5, 30, 10, 30, 13, 40, 50],
    "Week 1": [5, 8, 2, 50, 45, -15, 52, 20, 82],
    "Week 2": [7, 8, 1, 45, 15, -15, 38, 40, 80],
    "Week 3": [4, 8, -3, 15, 45, -15, 40, 40, 57],
    "Week 4": [11, 8, 0, 22, 18, -11, 38, 40, 55],
    "Week 5": [6, 8, -2, 15, 22, -18, 31, 33, 53],
    "Week 6": [10, 4, 4, 8, 28, -38, 20, 31, 28],
    "Week 7": [4, 4, 4, 10, 8, -36, 5, 6, 27],
    "Week 8": [4, 4, 4, 8, 8, -36, 4, 6, 25],
}

# Material cost for each product
material_cost = {
    "Lobster": 5,
    "Salmon": 2,
    "Shrimp": 3
}

df = pd.DataFrame(data)

# Inventory warning threshold (default value)
INVENTORY_WARNING_THRESHOLD = 150

def calculate_inventory(df):
    """Calculate weekly inventory based on the supply and demand formula, starting from Week 0"""
    week_columns = [col for col in df.columns if 'Week' in col]  # Get all week columns dynamically
    for week in range(0, len(week_columns)):  # Adjust the range to start from Week 0
        current_week = f"Week {week}"
        
        # Loop through products based on "Product" and "Item"
        products = df["Product"].unique()  # Ensure we work with unique products
        for product in products:
            try:
                supply_row = df.loc[(df["Product"] == product) & (df["Item"] == "Supply"), current_week].values[0]
                demand_row = df.loc[(df["Product"] == product) & (df["Item"] == "Demand"), current_week].values[0]

                if week == 0:  # For Week 0, there is no previous inventory, so we initialize directly
                    inventory_value = supply_row - demand_row
                else:
                    previous_week = f"Week {week-1}"
                    inventory_previous = df.loc[(df["Product"] == product) & (df["Item"] == "Inventory"), previous_week].values[0]
                    inventory_value = inventory_previous + supply_row - demand_row
                
                # Update inventory for the current week
                df.loc[(df["Product"] == product) & (df["Item"] == "Inventory"), current_week] = inventory_value
            except IndexError:
                continue
    return df

def update_table():
    """Update the table to display inventory data"""
    global df  # Ensure we use the global df
    df = calculate_inventory(df)
    for i in table.get_children():
        table.delete(i)

    # Insert each row and apply color formatting for inventory values
    for row in df.itertuples(index=False):
        row_data = list(row)
        item_id = table.insert("", "end", values=row_data)  # Insert the row without any tags
        if row_data[1] == "Inventory":  # Only for Inventory rows
            has_negative = any(value < 0 for value in row_data[2:])
            if has_negative:
                table.item(item_id, tags=("negative",))
            else:
                table.item(item_id, tags=("positive",))

def plot_inventory_value():
    """Plot the inventory value (inventory quantity * material cost)"""
    weeks = df.columns[2:]  # Dynamically get all week columns
    inventory_values = []

    # Calculate the inventory value for each week and each product
    for week in weeks:
        total_value = 0.0
        for product in df["Product"].unique():
            inventory_qty = df.loc[(df["Product"] == product) & (df["Item"] == "Inventory"), week].values[0]
            total_value += inventory_qty * material_cost[product]  # Dynamically apply the material cost
        inventory_values.append(total_value)
    
    # Clear previous plot
    for widget in graph_frame.winfo_children():
        widget.destroy()
    
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.plot(weeks, inventory_values, label='Inventory Value', marker='o')
    ax.axhline(y=INVENTORY_WARNING_THRESHOLD, color='r', linestyle='-', label='Warning Level')
    ax.set_title('Inventory Value vs Warning Level')
    ax.set_ylabel('Amount')
    ax.legend()
    
    canvas = FigureCanvasTkAgg(fig, master=graph_frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

    return fig  # Return the figure for saving

def modify_data():
    """Modify supply, demand, or unit cost based on user input"""
    product = product_choice_modify_week.get()
    week = week_choice_modify.get()
    
    try:
        supply_value = int(entry_modify_in.get())
        demand_value = int(entry_modify_out.get())
        new_cost = float(entry_modify_cost.get())
    except ValueError:
        messagebox.showerror("Error", "Please enter valid numbers for supply, demand, and cost.")
        return
    
    global df, material_cost  # Ensure we modify the global df and material cost

    # Modify supply and demand for the selected product and week
    df.loc[(df["Product"] == product) & (df["Item"] == "Supply"), week] = supply_value
    df.loc[(df["Product"] == product) & (df["Item"] == "Demand"), week] = demand_value

    # Modify material cost for the selected product
    material_cost[product] = new_cost  # Update material cost dynamically
    
    # Recalculate inventory based on the new supply and demand data
    df = calculate_inventory(df)  # Ensure we recalculate the inventory
    
    # Refresh the table and chart to reflect changes
    refresh_data()

    messagebox.showinfo("Update Successful", f"Updated supply, demand, and unit cost for {product} in {week}.")

def delete_product():
    """Delete a product from the dataframe"""
    product_to_delete = delete_product_combobox.get()

    if product_to_delete:
        global df, material_cost
        if product_to_delete in material_cost:
            # Remove the product from the material cost dictionary
            del material_cost[product_to_delete]

            # Remove all rows associated with the product from the dataframe
            df = df[df["Product"] != product_to_delete]
            
            refresh_data()
            refresh_product_list()  # Update product lists in comboboxes
            messagebox.showinfo("Success", f"Product '{product_to_delete}' deleted successfully.")
        else:
            messagebox.showerror("Error", f"Product '{product_to_delete}' not found.")
    else:
        messagebox.showerror("Error", "Please select a product.")

def refresh_product_list():
    """Refresh the product lists in the comboboxes"""
    products = list(df["Product"].unique())
    product_choice_modify_week['values'] = products
    delete_product_combobox['values'] = products

def modify_threshold():
    """Modify the inventory warning threshold based on user input"""
    try:
        new_threshold = int(entry_warning_threshold.get())
        global INVENTORY_WARNING_THRESHOLD  # 更新全局的警戒值
        INVENTORY_WARNING_THRESHOLD = new_threshold
        
        # 更新圖表，以反映新的警戒值
        refresh_data()

        messagebox.showinfo("Success", f"Inventory warning threshold set to {new_threshold}.")
    except ValueError:
        messagebox.showerror("Error", "Please enter a valid integer for the warning threshold.")

def add_product():
    """Add a new product to the dataframe"""
    new_product = entry_add_product.get()
    new_cost = entry_add_cost.get()

    if new_product and new_cost:
        try:
            new_cost = float(new_cost)
            global df, material_cost
            
            # Add the new product to material_cost
            material_cost[new_product] = new_cost

            # Create default Supply, Demand, Inventory for the new product
            supply_row = [new_product, "Supply"] + [0] * (df.shape[1] - 2)
            demand_row = [new_product, "Demand"] + [0] * (df.shape[1] - 2)
            inventory_row = [new_product, "Inventory"] + [0] * (df.shape[1] - 2)

            # Append the new rows to the dataframe
            new_rows = pd.DataFrame([supply_row, demand_row, inventory_row], columns=df.columns)
            df = pd.concat([df, new_rows], ignore_index=True)
            
            refresh_data()
            refresh_product_list()  # Update product lists in comboboxes
            messagebox.showinfo("Success", f"Product '{new_product}' added successfully.")
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid cost.")
    else:
        messagebox.showerror("Error", "Please enter both product name and cost.")

def refresh_data():
    """Refresh the table and plot to reflect the latest data"""
    update_table()  # Updates the table with the modified inventory data
    plot_inventory_value()  # Re-plots the inventory value graph

def save_chart_as_png():
    """Save the current chart as a PNG file"""
    fig = plot_inventory_value()  # Get the current figure
    file_path = tk.filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
    if file_path:
        fig.savefig(file_path, format="png")
        messagebox.showinfo("Success", f"Chart saved as {file_path}")

def save_table_as_excel():
    """Save the current table as an Excel file"""
    file_path = tk.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Table saved as {file_path}")

# Create main window
root = tk.Tk()
root.title("Inventory Management System")
root.geometry("1400x1000")

# Create input area for modifying supply, demand, and unit cost in one action
frame_input_modify = tk.Frame(root)
frame_input_modify.grid(row=0, column=0, padx=10, pady=10, sticky="nw")

tk.Label(frame_input_modify, text="Select Product:").grid(row=0, column=0)
product_choice_modify_week = ttk.Combobox(frame_input_modify, values=list(df["Product"].unique()))
product_choice_modify_week.grid(row=0, column=1)

tk.Label(frame_input_modify, text="Select Week:").grid(row=1, column=0)
week_choice_modify = ttk.Combobox(frame_input_modify, values=[f"Week {i}" for i in range(0, 9)])
week_choice_modify.grid(row=1, column=1)

tk.Label(frame_input_modify, text="Modify Supply:").grid(row=2, column=0)
entry_modify_in = tk.Entry(frame_input_modify)
entry_modify_in.grid(row=2, column=1)

tk.Label(frame_input_modify, text="Modify Demand:").grid(row=3, column=0)
entry_modify_out = tk.Entry(frame_input_modify)
entry_modify_out.grid(row=3, column=1)

tk.Label(frame_input_modify, text="Modify Unit Cost:").grid(row=4, column=0)
entry_modify_cost = tk.Entry(frame_input_modify)
entry_modify_cost.grid(row=4, column=1)

btn_modify_data = tk.Button(frame_input_modify, text="Modify Data", command=modify_data)
btn_modify_data.grid(row=5, columnspan=2, pady=10)

# Create input area for modifying inventory warning threshold
frame_threshold_modify = tk.Frame(root)
frame_threshold_modify.grid(row=0, column=1, padx=10, pady=10, sticky="n")  # Center-aligned

tk.Label(frame_threshold_modify, text="Set Inventory Warning Threshold:").grid(row=0, column=0)
entry_warning_threshold = tk.Entry(frame_threshold_modify)
entry_warning_threshold.grid(row=0, column=1)

btn_modify_threshold = tk.Button(frame_threshold_modify, text="Set Threshold", command=modify_threshold)
btn_modify_threshold.grid(row=1, columnspan=2, pady=10)

# Create input area for adding/deleting products
frame_product_modify = tk.Frame(root)
frame_product_modify.grid(row=0, column=2, padx=10, pady=10, sticky="n")  # Right-aligned

tk.Label(frame_product_modify, text="Add Product:").grid(row=0, column=0)
entry_add_product = tk.Entry(frame_product_modify)
entry_add_product.grid(row=0, column=1)

tk.Label(frame_product_modify, text="Cost:").grid(row=1, column=0)
entry_add_cost = tk.Entry(frame_product_modify)
entry_add_cost.grid(row=1, column=1)

btn_add_product = tk.Button(frame_product_modify, text="Add Product", command=add_product)
btn_add_product.grid(row=2, columnspan=2, pady=10)

# Create buttons for saving chart as PNG and table as Excel
frame_save_buttons = tk.Frame(root)
frame_save_buttons.grid(row=4, column=0, padx=10, pady=10, sticky="s")

btn_save_chart = tk.Button(frame_save_buttons, text="Save Chart as PNG", command=save_chart_as_png)
btn_save_chart.pack(side=tk.LEFT, padx=10)

btn_save_table = tk.Button(frame_save_buttons, text="Save Table as Excel", command=save_table_as_excel)
btn_save_table.pack(side=tk.LEFT, padx=10)

tk.Label(frame_product_modify, text="Delete Product:").grid(row=3, column=0)
delete_product_combobox = ttk.Combobox(frame_product_modify, values=list(df["Product"].unique()))
delete_product_combobox.grid(row=3, column=1)

btn_delete_product = tk.Button(frame_product_modify, text="Delete Product", command=delete_product)
btn_delete_product.grid(row=4, columnspan=2, pady=10)

# Create table area (top-left)
frame_table = tk.Frame(root)
frame_table.grid(row=1, column=0, padx=10, pady=10, sticky="nw", columnspan=2)

columns = list(df.columns[:2]) + list(df.columns[2:])  # Keep columns as is
table = ttk.Treeview(frame_table, columns=columns, show="headings", height=12)
table.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

# Set table column headings
for col in columns:
    table.heading(col, text=col)

# Set column width
for col in columns:
    table.column(col, width=100, anchor="center")

# Define tags for positive (blue) and negative (red) values
table.tag_configure('positive', foreground='blue')
table.tag_configure('negative', foreground='red')

# Create graph display area (bottom-left)
graph_frame = tk.Frame(root)
graph_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nw", columnspan=2)

# Create summary report area (bottom, spanning both columns)
summary_frame = tk.Frame(root)
summary_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="s")

# Initial table and graph update
refresh_data()

# Start main loop
root.mainloop()

