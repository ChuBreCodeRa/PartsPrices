import pandas as pd
import re
import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinter import ttk
import tkinter.font as tkFont
import os
import datetime

def extract_part_numbers(input_text):
    return re.findall(r'\b\d{6}\b|\b\d{10}\b', input_text)

def find_price_list_file(base_path, file_prefix):
    """Find the latest file in the directory that starts with the given prefix."""
    try:
        files = [f for f in os.listdir(base_path) if f.startswith(file_prefix) and (f.endswith(".xlsx") or f.endswith(".xls"))]
        if not files:
            messagebox.showerror("File Not Found", f"No file starting with '{file_prefix}' was found in '{base_path}'.")
            return None
        # Sort by last modified time and pick the latest file
        files = sorted(files, key=lambda f: os.path.getmtime(os.path.join(base_path, f)), reverse=True)
        return os.path.join(base_path, files[0])
    except Exception as e:
        messagebox.showerror("Error", f"Error accessing the directory: {e}")
        return None

def load_price_list(region, pricing_type):
    base_path = r"S:\Consumer Services\Consumer Service SOP (WIP)\Dealer Support\Parts Price List"
    file_prefix = "CNA_Parts_Price_List_All"
    price_list_path = find_price_list_file(base_path, file_prefix)

    if price_list_path is None:
        return None

    try:
        # Determine the appropriate engine based on the file extension
        if price_list_path.endswith(".xls"):
            return pd.read_excel(price_list_path, engine='xlrd')
        elif price_list_path.endswith(".xlsx"):
            return pd.read_excel(price_list_path, engine='openpyxl')
        else:
            messagebox.showerror("File Error", "Unsupported file format for price list.")
            return None
    except FileNotFoundError:
        messagebox.showerror("File Not Found", f"The file at '{price_list_path}' was not found.")
    except Exception as e:
        messagebox.showerror("Error", f"Error reading the price list: {e}")
    return None

def load_inventory():
    inventory_file = r"S:\Consumer Services\Consumer Service SOP (WIP)\Dealer Support\Parts Help Desk\CNA Inventory QTY.xlsx"
    try:
        # Check last modified date
        last_modified_timestamp = os.path.getmtime(inventory_file)
        last_modified_date = datetime.datetime.fromtimestamp(last_modified_timestamp)
        current_date = datetime.datetime.now()
        days_since_modified = (current_date - last_modified_date).days

        if days_since_modified > 7:
            messagebox.showwarning("Inventory File Outdated", "The inventory file is more than a week old. Please update it from the inventory dashboard.")

        # Update the label with the last modified date
        last_modified_label_var.set(f"Inventory Last Modified: {last_modified_date.strftime('%Y-%m-%d %H:%M:%S')}")

        return pd.read_excel(inventory_file, engine='openpyxl')
    except FileNotFoundError:
        messagebox.showerror("File Not Found", "The file 'CNA Inventory QTY.xlsx' was not found.")
    except Exception as e:
        messagebox.showerror("Error", f"Error reading the Excel file: {e}")
    return None

def extract_inventory_levels(df):
    """
    For each part number, go over and the right 1 column and scan downward in column index 3
    until "Item Total" is found. Then, the value in column index 10 is used as that part's
    inventory level.
    
    Returns a dictionary mapping each part number (as a string) to its inventory level.
    """
    inventory_levels = {}
    n_rows = df.shape[0]
    
    for i in range(n_rows):
        part_number = df.iloc[i, 2]
        if pd.notnull(part_number) and str(part_number).strip() != "":
            for j in range(i + 1, n_rows):
                cell_val = df.iloc[j, 3]
                if isinstance(cell_val, str) and "Item Total" in cell_val:
                    inv_value = df.iloc[j, 10]
                    try:
                        inv_level = float(inv_value)
                    except (ValueError, TypeError):
                        inv_level = 0.0
                    inventory_levels[str(part_number).strip()] = inv_level
                    break  
    return inventory_levels

def search_parts():
    region = region_var.get().lower()
    pricing_type = pricing_type_var.get().lower()
    input_text = part_numbers_input.get("1.0", tk.END)
    part_numbers_from_input = extract_part_numbers(input_text)

    # If no part numbers were entered, show a warning and return
    if not part_numbers_from_input:
        messagebox.showwarning("No Parts Entered", "Please enter at least one part number.")
        return
    
    # Load the price list and inventory data
    df_price_list = load_price_list(region, pricing_type)
    df_inventory = load_inventory()
    
    # If either of the dataframes is None, return without further processing
    if df_price_list is None or df_inventory is None:
        return

    # Build a dictionary from the price list. Part numbers in column index 2, descriptions in column index 4, and prices in column index 5.
    part_numbers = df_price_list.iloc[:, 2]
    descriptions = df_price_list.iloc[:, 4]
    prices = df_price_list.iloc[:, 5]
    part_info_dict = {
        str(part_number).strip(): (description, price)
        for part_number, description, price in zip(part_numbers, descriptions, prices)
    }

    # Build a dictionary from the inventory data.
    inventory_dict = extract_inventory_levels(df_inventory)

    # Prepare the output text
    output_text = "{:<15} {:<30} {:<15} {:<15}\n".format("Part #", "Description", "Cost", "Availability")
    
    # Build a separate string for stock levels (numeric values)
    stock_info_text = "Stock Levels:\n"
    
    for part_number in part_numbers_from_input:
        part_number = part_number.strip()
        if part_number in part_info_dict:
            description, price = part_info_dict[part_number]
            
            # Determine the inventory level and corresponding availability.
            inv_level = inventory_dict.get(part_number, None)
            if inv_level is None:
                availability_status = "N/A"
                stock_value = "Not Found"
            else:
                stock_value = inv_level 
                if inv_level >= 10:
                    availability_status = "Healthy stock"
                elif inv_level > 0:
                    availability_status = "Low stock"
                else:
                    availability_status = "Out of stock"

            if isinstance(price, str):
                try:
                    price = float(price.replace('$', '').replace(',', ''))
                except ValueError:
                    price = 0.0

            formatted_price = f"${price:,.2f}"
            output_text += "{:<15} {:<30} {:<15} {:<15}\n".format(part_number, description, formatted_price, availability_status)
            stock_info_text += f"{part_number}: {stock_value}\n"
        else:
            output_text += "{:<15} {:<30} {:<15} {:<15}\n".format(part_number, "Not found in price list", "N/A", "N/A")
            stock_info_text += f"{part_number}: Not Found\n"

    output_display.delete("1.0", tk.END)
    output_display.insert(tk.END, output_text)
    
    # Update the non-copyable stock information label
    stock_info_var.set(stock_info_text)

# GUI Setup
root = tk.Tk()
root.title("Part Inventory Lookup")
root.geometry("800x650") 

# Set a monospaced font for better alignment
monospace_font = tkFont.Font(family="Courier", size=10)

# Region Selection
region_label = tk.Label(root, text="Select Region:")
region_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
region_var = tk.StringVar(value="US")
region_menu = ttk.Combobox(root, textvariable=region_var, values=["Canada", "US"], state="readonly")
region_menu.grid(row=0, column=1, padx=10, pady=10, sticky="w")

# Pricing Type Selection
pricing_type_label = tk.Label(root, text="Select Pricing Type:")
pricing_type_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
pricing_type_var = tk.StringVar(value="Dealer")
pricing_type_menu = ttk.Combobox(root, textvariable=pricing_type_var, values=["Dealer", "Distributor"], state="readonly")
pricing_type_menu.grid(row=1, column=1, padx=10, pady=10, sticky="w")

# Part Numbers Input
part_numbers_label = tk.Label(root, text="Enter Part Numbers:")
part_numbers_label.grid(row=2, column=0, padx=10, pady=10, sticky="nw")
part_numbers_input = scrolledtext.ScrolledText(root, width=50, height=10, wrap=tk.WORD)
part_numbers_input.grid(row=2, column=1, padx=10, pady=10)

# Search Button
search_button = tk.Button(root, text="Search", command=search_parts)
search_button.grid(row=3, column=1, padx=10, pady=10, sticky="e")

# Last Modified Date Label
last_modified_label_var = tk.StringVar(value="Inventory Last Modified: Not checked yet")
last_modified_label = tk.Label(root, textvariable=last_modified_label_var)
last_modified_label.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="w")

# Output Display (copyable text area)
output_display = scrolledtext.ScrolledText(root, width=100, height=15, font=monospace_font)
output_display.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

# Stock Info Display (non-copyable - using a Label)
stock_info_var = tk.StringVar(value="Stock Levels: (Not searched yet)")
stock_info_label = tk.Label(root, textvariable=stock_info_var, anchor="w", justify="left", bg="lightgrey", font=("Courier", 10))
stock_info_label.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="w")

root.mainloop()
