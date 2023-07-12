# Import necessary modules
import tkinter as tk  # GUI toolkit
import sqlite3  # SQLite database
import pandas as pd  # Data manipulation and analysis library

# Establish a connection to the database
conn = sqlite3.connect('inventory.db')
cursor = conn.cursor()

# Create a table to store products if it doesn't already exist
cursor.execute('''CREATE TABLE IF NOT EXISTS Products (
                    product_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    description TEXT,
                    quantity INTEGER,
                    unit_price REAL,
                    supplier TEXT
                )''')

# Function to add a new product to the inventory
def add_product():
    add_window = tk.Toplevel(root)
    add_window.title("Add Product")

    # Create labels and entry fields for product information
    tk.Label(add_window, text="Product Name:").grid(row=0, column=0)
    name_entry = tk.Entry(add_window)
    name_entry.grid(row=0, column=1)

    tk.Label(add_window, text="Product Description:").grid(row=1, column=0)
    description_entry = tk.Entry(add_window)
    description_entry.grid(row=1, column=1)

    tk.Label(add_window, text="Product Quantity:").grid(row=2, column=0)
    quantity_entry = tk.Entry(add_window)
    quantity_entry.grid(row=2, column=1)

    tk.Label(add_window, text="Product Unit Price:").grid(row=3, column=0)
    unit_price_entry = tk.Entry(add_window)
    unit_price_entry.grid(row=3, column=1)

    tk.Label(add_window, text="Product Supplier:").grid(row=4, column=0)
    supplier_entry = tk.Entry(add_window)
    supplier_entry.grid(row=4, column=1)

    def add_product():
        # Retrieve the product information entered by the user
        name = name_entry.get()
        description = description_entry.get()
        quantity = int(quantity_entry.get())
        unit_price = float(unit_price_entry.get())
        supplier = supplier_entry.get()

        # Insert the product into the database
        cursor.execute('''INSERT INTO Products (name, description, quantity, unit_price, supplier)
                          VALUES (?, ?, ?, ?, ?)''', (name, description, quantity, unit_price, supplier))
        conn.commit()
        status_label['text'] = "Product added successfully."
        add_window.destroy()

    add_button = tk.Button(add_window, text="Add Product", command=add_product)
    add_button.grid(row=5, column=0, columnspan=2, pady=5)

# Function to update the quantity, unit price, or supplier of a product
def update_product():
    update_window = tk.Toplevel(root)
    update_window.title("Update Product")

    # Create labels and entry fields for product ID, field to update, and new value
    tk.Label(update_window, text="Product ID:").grid(row=0, column=0)
    product_id_entry = tk.Entry(update_window)
    product_id_entry.grid(row=0, column=1)

    tk.Label(update_window, text="Field to Update:").grid(row=1, column=0)
    field_entry = tk.Entry(update_window)
    field_entry.grid(row=1, column=1)

    tk.Label(update_window, text="New Value:").grid(row=2, column=0)
    new_value_entry = tk.Entry(update_window)
    new_value_entry.grid(row=2, column=1)

    status_label = tk.Label(root, text="")

    def update_product():
        # Retrieve the product ID, field, and new value entered by the user
        product_id = int(product_id_entry.get())
        field = field_entry.get()
        new_value = new_value_entry.get()

        # Update the corresponding field of the product in the database
        if field == "quantity":
            new_quantity = int(new_value)
            cursor.execute('''UPDATE Products SET quantity = ? WHERE product_id = ?''', (new_quantity, product_id))
            conn.commit()
            status_label['text'] = "Quantity updated successfully."
        elif field == "unit_price":
            new_unit_price = float(new_value)
            cursor.execute('''UPDATE Products SET unit_price = ? WHERE product_id = ?''', (new_unit_price, product_id))
            conn.commit()
            status_label['text'] = "Unit price updated successfully."
        elif field == "supplier":
            cursor.execute('''UPDATE Products SET supplier = ? WHERE product_id = ?''', (new_value, product_id))
            conn.commit()
            status_label['text'] = "Supplier updated successfully."
        else:
            status_label['text'] = "Invalid field. Please try again."

        update_window.destroy()

    update_button = tk.Button(update_window, text="Update Product", command=update_product)
    update_button.grid(row=3, column=0, columnspan=2, pady=5)

# Function to delete a product from the inventory
def delete_product():
    delete_window = tk.Toplevel(root)
    delete_window.title("Delete Product")

    # Create labels and entry fields for product ID
    tk.Label(delete_window, text="Product ID:").grid(row=0, column=0)
    product_id_entry = tk.Entry(delete_window)
    product_id_entry.grid(row=0, column=1)

    status_label = tk.Label(root, text="")

    def delete_product():
        # Retrieve the product ID entered by the user
        product_id = int(product_id_entry.get())

        # Delete the product from the database
        cursor.execute('''DELETE FROM Products WHERE product_id = ?''', (product_id,))
        conn.commit()
        status_label['text'] = "Product deleted successfully."
        delete_window.destroy()

    delete_button = tk.Button(delete_window, text="Delete Product", command=delete_product)
    delete_button.grid(row=1, column=0, columnspan=2, pady=5)

# Function to list all products in the inventory
def list_products():
    list_window = tk.Toplevel(root)
    list_window.title("List Products")

    # Retrieve all products from the database
    cursor.execute('''SELECT * FROM Products''')
    products = cursor.fetchall()

    output_textbox = tk.Text(list_window, width=50, height=10)
    output_textbox.pack(pady=10)

    # Display the product information in the text box
    for product in products:
        output_textbox.insert(tk.END, f"Product ID: {product[0]}\n")
        output_textbox.insert(tk.END, f"Name: {product[1]}\n")
        output_textbox.insert(tk.END, f"Description: {product[2]}\n")
        output_textbox.insert(tk.END, f"Quantity: {product[3]}\n")
        output_textbox.insert(tk.END, f"Unit Price: ${product[4]}\n")
        output_textbox.insert(tk.END, f"Supplier: {product[5]}\n")
        output_textbox.insert(tk.END, "\n")

    close_button = tk.Button(list_window, text="Close", command=list_window.destroy)
    close_button.pack(pady=10)

# Function to import products from an Excel file
def import_products():
    df = pd.read_excel('products.xlsx')

    # Iterate over each row in the Excel file and insert the products into the database
    for index, row in df.iterrows():
        name = row['Name']
        description = row['Description']
        quantity = row['Quantity']
        unit_price = row['Unit Price($)']
        supplier = row['Supplier']

        # Insert the product into the database
        cursor.execute('''INSERT INTO Products (name, description, quantity, unit_price, supplier)
                          VALUES (?, ?, ?, ?, ?)''', (name, description, quantity, unit_price, supplier))
        conn.commit()

    status_label['text'] = "Products imported successfully."

# Function to export products to an Excel file
def export_products():
    # Retrieve all products from the database
    cursor.execute('''SELECT * FROM Products''')
    products = cursor.fetchall()

    # Create a DataFrame from the products and export it to an Excel file
    df = pd.DataFrame(products, columns=['Product ID', 'Name', 'Description', 'Quantity', 'Unit Price($)', 'Supplier'])
    df.to_excel('products_export.xlsx', index=False)

    status_label['text'] = "Products exported successfully."

# Function to display the help message
def display_help():
    help_window = tk.Toplevel(root)
    help_window.title("Help")

    help_text = """
    Inventory Help
    ---------------------
    Commands:
    Add Product - Add a new product to the inventory
    Update Product - Update the quantity of a product
    Delete Product - Delete a product from the inventory
    List Products - Retrieve all products from the inventory
    Import Products - Import products from an Excel file
    Export Products - Export products to an Excel file
    Exit - Exit the program
    Help - Display this help message
    """

    tk.Label(help_window, text=help_text).pack(padx=10, pady=10)

    close_button = tk.Button(help_window, text="Close", command=help_window.destroy)
    close_button.pack(pady=10)

# Create the main window
root = tk.Tk()
root.title("Inventory Management System")

# Create buttons for various actions
add_button = tk.Button(root, text="Add Product", command=add_product)
add_button.grid(row=0, column=0, padx=10, pady=10)

update_button = tk.Button(root, text="Update Product", command=update_product)
update_button.grid(row=0, column=1, padx=10, pady=10)

delete_button = tk.Button(root, text="Delete Product", command=delete_product)
delete_button.grid(row=1, column=0, padx=10, pady=10)

list_button = tk.Button(root, text="List Products", command=list_products)
list_button.grid(row=1, column=1, padx=10, pady=10)

import_button = tk.Button(root, text="Import Products", command=import_products)
import_button.grid(row=2, column=0, padx=10, pady=10)

export_button = tk.Button(root, text="Export Products", command=export_products)
export_button.grid(row=2, column=1, padx=10, pady=10)

help_button = tk.Button(root, text="Help", command=display_help)
help_button.grid(row=3, column=0, padx=10, pady=10)

exit_button = tk.Button(root, text="Exit", command=root.destroy)
exit_button.grid(row=3, column=1, padx=10, pady=10)

status_label = tk.Label(root, text="")
status_label.grid(row=4, column=0, columnspan=2)

root.mainloop()

# Close the database connection
conn.close()
