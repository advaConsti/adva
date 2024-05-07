import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog
from PIL import Image, ImageTk
import os
import sys

# Scrollable Frame Class
class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def parse_xml_file(file_path):
    tree = ET.parse(file_path)
    return tree.getroot()

def parse_customers(root):
    customers = []
    for customer in root.findall('.//Customer'):
        SellToCustomer = customer.find('No').text if customer.find('No') is not None else ''
        name = customer.find('Name').text if customer.find('Name') is not None else ''
        address = customer.find('Address').text if customer.find('Address') is not None else ''
        postcode = customer.find('PostCode').text if customer.find('PostCode') is not None else ''
        city = customer.find('City').text if customer.find('City') is not None else ''
        paymentmethod = customer.find('PaymentMethodCode').text if customer.find('PaymentMethodCode') is not None else ''
        paymenttime = customer.find('PaymentTermsCode').text if customer.find('PaymentTermsCode') is not None else ''
        customers.append([SellToCustomer, name, address, postcode, city, paymentmethod, paymenttime])
    return pd.DataFrame(customers, columns=['No', 'Name', 'Address', 'PostCode', 'City', 'Payment Method', 'Payment time frame'])

def parse_invoices(root, customers_df):
    invoices_data = []
    unique_invoice_nos = set()

    for invoice in root.findall('.//Invoice'):
        invoice_no = invoice.find('No').text if invoice.find('No') is not None else ''
        unique_invoice_nos.add(invoice_no)

    for invoice_no in unique_invoice_nos:
        for invoice in root.findall(".//Invoice[No='" + invoice_no + "']"):
            sell_to_customer_no = invoice.find('SellToCustomer').text if invoice.find('SellToCustomer') is not None else ''
            customer_data = customers_df[customers_df['No'] == sell_to_customer_no].iloc[0] if not customers_df[customers_df['No'] == sell_to_customer_no].empty else {}

            invoice_totals = {
                'InvoiceNo': invoice_no,
                'DEB-No': sell_to_customer_no,
                'CustomerName': customer_data.get('Name', '') if not customer_data.empty else '',
                'Fulfillment-Dienstleistung': 0,
                'Versandkosten': 0
            }

            for line in invoice.findall('.//Line'):
                description = line.find('Description').text if line.find('Description') is not None else ''
                sales_price_text = line.find('SalesPrice').text if line.find('SalesPrice') is not None else '0'
                sales_price_text = sales_price_text.replace('.', '').replace(',', '.')

                try:
                    sales_price = float(sales_price_text)
                except ValueError:
                    sales_price = 0

                if description == 'Fulfillment-Dienstleistung':
                    invoice_totals['Fulfillment-Dienstleistung'] += sales_price
                elif description in ['Versandkosten EU', 'Versandkosten Non-EU']:
                    invoice_totals['Versandkosten'] += sales_price

            invoices_data.append(invoice_totals)

    return pd.DataFrame(invoices_data)

def display_customers_and_invoices(root, customers_df, invoices_df):
    global selected_customers
    frame = ScrollableFrame(root)
    frame.place(relx=0.5, rely=0.7, anchor='center', relwidth=0.9, relheight=0.4)

    headers = ["Select", "Customer No", "Name", "Invoice No", "Fulfillment-Dienstleistung", "Versandkosten"]
    for col, header in enumerate(headers):
        header_label = ttk.Label(frame.scrollable_frame, text=header, font=("Arial", 10, "bold"))
        header_label.grid(row=0, column=col, padx=5, pady=5)

    selected_customers = {str(index): tk.BooleanVar() for index in customers_df.index}

    for index, row in customers_df.iterrows():
        checkbox = ttk.Checkbutton(frame.scrollable_frame, variable=selected_customers[str(index)])
        checkbox.grid(row=index + 1, column=0)

        ttk.Label(frame.scrollable_frame, text=row['No']).grid(row=index + 1, column=1)
        ttk.Label(frame.scrollable_frame, text=row['Name']).grid(row=index + 1, column=2)

        customer_invoices = invoices_df[invoices_df['DEB-No'] == row['No']]
        invoice_nos = ', '.join(customer_invoices['InvoiceNo'].unique())
        total_fulfillment = customer_invoices['Fulfillment-Dienstleistung'].sum()
        total_shipping = customer_invoices['Versandkosten'].sum()

        ttk.Label(frame.scrollable_frame, text=invoice_nos).grid(row=index + 1, column=3)
        ttk.Label(frame.scrollable_frame, text=f"{total_fulfillment:.2f}").grid(row=index + 1, column=4)
        ttk.Label(frame.scrollable_frame, text=f"{total_shipping:.2f}").grid(row=index + 1, column=5)

def load_xml_file():
    filename = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
    if filename:
        global xml_file_path
        xml_file_path = filename
        xml_entry.delete(0, tk.END)
        xml_entry.insert(0, filename)
        status_label.config(text="Selected XML: " + filename)

def select_output_folder():
    foldername = filedialog.askdirectory()
    if foldername:
        global output_folder
        output_folder = foldername
        output_entry.delete(0, tk.END)
        output_entry.insert(0, foldername)
        status_label.config(text="Selected Output Folder: " + foldername)

def start_processing():
    if xml_file_path and output_folder:
        global root_element, customers_df, invoices_df
        root_element = parse_xml_file(xml_file_path)
        customers_df = parse_customers(root_element)
        invoices_df = parse_invoices(root_element, customers_df)

        display_customers_and_invoices(root, customers_df, invoices_df)

        status_label.config(text="Display Complete")
    else:
        status_label.config(text="Please select XML and output folder")

def save_filtered_xml():
    if not customers_df.empty:
        # Create a new XML root element using the original root's tag and attributes
        new_root = ET.Element(root_element.tag, root_element.attrib)

        # Get customer numbers that are marked for removal
        customers_to_remove = [
            customers_df.loc[int(index), 'No'] for index, selected in selected_customers.items() if selected.get()
        ]

        print("Customers to remove:", customers_to_remove)  # Debugging statement

        # Filtering and adding customers who are not marked for removal
        for customer in root_element.findall('.//Customer'):
            customer_no = customer.find('No').text
            if customer_no not in customers_to_remove:
                new_root.append(customer)
            else:
                print("Removing customer:", customer_no)  # Debugging statement

        # Add invoices only if their corresponding customer is not marked for removal
        for invoice in root_element.findall('.//Invoice'):
            sell_to_customer_no = invoice.find('SellToCustomer').text
            print("Invoice checked for:", sell_to_customer_no)  # Debugging statement
            if sell_to_customer_no not in customers_to_remove:
                new_root.append(invoice)
            else:
                print("Removing invoice for customer:", sell_to_customer_no)  # Debugging statement

        # Write the new XML tree to a file
        output_file = os.path.join(output_folder, f"{xml_file_path}")
        new_tree = ET.ElementTree(new_root)
        new_tree.write(output_file, encoding='utf-8', xml_declaration=True)

        status_label.config(text=f"Filtered XML saved to {output_file}")



# GUI setup
root = tk.Tk()
root.title("XML to Excel Converter")
image_path = resource_path("VB_1.png")
img = Image.open(image_path)
window_width, window_height = img.size
root.geometry(f'{window_width}x{window_height}')
background_image = ImageTk.PhotoImage(img)

background_label = tk.Label(root, image=background_image)
background_label.place(relwidth=1, relheight=1)

xml_file_path = ""
output_folder = ""

xml_entry = tk.Entry(root, width=40)
xml_entry.place(relx=0.5, rely=0.1, anchor='center')

xml_browse_button = tk.Button(root, text="Browse XML", command=load_xml_file, bg='#80c1ff')
xml_browse_button.place(relx=0.5, rely=0.15, anchor='center')

output_entry = tk.Entry(root, width=40)
output_entry.place(relx=0.5, rely=0.2, anchor='center')

output_browse_button = tk.Button(root, text="Select Output Folder", command=select_output_folder, bg='#80c1ff')
output_browse_button.place(relx=0.5, rely=0.25, anchor='center')

process_button = tk.Button(root, text="Start Processing", command=start_processing, bg='#80c1ff')
process_button.place(relx=0.5, rely=0.3, anchor='center')

save_button = tk.Button(root, text="Save Filtered XML", command=save_filtered_xml, bg='#80c1ff')
save_button.place(relx=0.5, rely=0.35, anchor='center')

status_label = tk.Label(root, text="", fg="green")
status_label.place(relx=0.5, rely=0.4, anchor='center')

root.mainloop()
