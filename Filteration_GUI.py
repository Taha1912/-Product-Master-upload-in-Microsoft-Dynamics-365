#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import tkinter as tk
from tkinter import filedialog, messagebox, Label, IntVar, Checkbutton
import pandas as pd
import re

output_files = [
        "01-PRODUCTS V2.xlsx", "02-RELEASED PRODUCTS V2(J.).xlsx","02-RELEASED PRODUCTS V2(UI).xlsx", "03-PRODUCT MASTER COLOR.xlsx", 
        "04-PRODUCT MASTER SIZE.xlsx", "05-PRODUCT MASTER STYLE.xlsx", "06-RELEASED PRODUCT VARIANTS V2.xlsx",
        "07-ITEM BARCODE.xlsx", "08-PRODUCT CATEGORY ASSIGNMENTS.xlsx", "09-ITEM BATCH.xlsx", "COLOR.xlsx", "SIZE.xlsx", "STYLE.xlsx"]
    

def transfer_data(input_file, output_folder, default_barcodesetupid, default_productcategoryname, apply_custom_transformation, selected_files):
    global output_files  # Access the global output_files list
    
    # Determine which files are selected based on the values in selected_files
    selected_output_files = [output_files[i] for i, selected in enumerate(selected_files) if selected.get() == 1]
    
    # Load the Excel file into a DataFrame
    df_input = pd.read_excel(input_file, dtype=str, keep_default_na=False)
    
    for output_file in selected_output_files:  # Iterate through selected output files

        # Load the output Excel file into a DataFrame
        df_output = pd.read_excel(output_file)

        # Initialize a DataFrame to hold the data to be transferred
        df_transfer = pd.DataFrame()
        max_length = 20
        
        # Conditions to match columns and transfer data
        for column_name in df_output.columns:
            if column_name == "BARCODESETUPID":
                if default_barcodesetupid.get() == 1:
                    df_transfer[column_name] = "JDOT"
                elif default_barcodesetupid.get() == 2:
                    df_transfer[column_name] = "ALM"
                elif default_barcodesetupid.get() == 3:
                    df_transfer[column_name] = "UI"

            elif column_name == "PRODUCTCATEGORYHIERARCHYNAME":
                if default_productcategoryname.get() == 1:
                    df_transfer[column_name] = "RETAIL CATEGORY"
                elif default_productcategoryname.get() == 2:
                    df_transfer[column_name] = "PROCUREMENT CATEGORY"
                    
            elif column_name == "UNITCOST":
                df_transfer[column_name] = df_input["PURCHASEPRICE"]
                
            
            elif column_name == "BATCHNUMBER":
                df_transfer[column_name] = df_input["BATCH NAME"]
            elif column_name == "BATCHDESCRIPTION":
                df_transfer[column_name] = df_input["BATCH NAME"]
                    
                    
            # Custom transformation for UNITCOST column if selected by the user
            elif column_name == "PURCHASEUNDERDELIVERYPERCENTAGE" and apply_custom_transformation.get():
                df_transfer[column_name] = (df_input["PURCHASEUNDERDELIVERYPERCENTAGE"].astype(float) * 100)
                
            elif column_name == "PURCHASEOVERDELIVERYPERCENTAGE" and apply_custom_transformation.get():
                df_transfer[column_name] = (df_input["PURCHASEOVERDELIVERYPERCENTAGE"].astype(float) * 100)
                
            elif column_name == "SALESUNDERDELIVERYPERCENTAGE" and apply_custom_transformation.get():
                df_transfer[column_name] = (df_input["SALESUNDERDELIVERYPERCENTAGE"].astype(float) * 100)
                
            elif column_name == "SALESOVERDELIVERYPERCENTAGE" and apply_custom_transformation.get():
                df_transfer[column_name] = (df_input["SALESOVERDELIVERYPERCENTAGE"].astype(float) * 100)
                
            elif column_name == "TRANSFERORDERUNDERDELIVERYPERCENTAGE" and apply_custom_transformation.get():
                df_transfer[column_name] = (df_input["TRANSFERORDERUNDERDELIVERYPERCENTAGE"].astype(float) * 100)
                
            elif column_name == "TRANSFERORDEROVERDELIVERYPERCENTAGE" and apply_custom_transformation.get():
                df_transfer[column_name] = (df_input["TRANSFERORDEROVERDELIVERYPERCENTAGE"].astype(float) * 100)
                
                
            # Automatically match columns and transfer data
            elif column_name in df_input.columns:
                df_transfer[column_name] = df_input[column_name]
                
            # Predefined match columns and transfer data

            elif column_name == "PRODUCTSEARCHNAME":
                df_transfer[column_name] = df_input["PRODUCTDESCRIPTION"].str[:max_length]
            elif column_name == "PRODUCTNUMBER":
                df_transfer[column_name] = df_input["ITEMNUMBER"]
            elif column_name == "SEARCHNAME":
                df_transfer[column_name] = df_input["PRODUCTDESCRIPTION"]
            elif column_name == "BUYERGROUPID":
                df_transfer[column_name] = df_input["PCTCodes"]
            elif column_name == "PRODUCTMASTERNUMBER":
                df_transfer[column_name] = df_input["ITEMNUMBER"]
            elif column_name == "TRANSLATEDCOLORDESCRIPTION":
                df_transfer[column_name] = df_input["COLORDESCRIPTION"]
            elif column_name == "TRANSLATEDSIZEDESCRIPTION":
                df_transfer[column_name] = df_input["SIZEDESCRIPTION"]
            elif column_name == "TRANSLATEDSTYLEDESCRIPTION":
                df_transfer[column_name] = df_input["STYLEDESCRIPTION"]
            elif column_name == "PRODUCTCATEGORYNAME":
                df_transfer[column_name] = df_input["RETAILPRODUCTCATEGORYNAME.6"]
                
            elif column_name == "DEFAULTLEDGERDIMENSIONDISPLAYVALUE":
                df_transfer[column_name] = "----" + (df_input["PRODUCTGROUPID"].astype(str)).apply(lambda x: re.search(r'\d+', str(x)).group() if re.search(r'\d+', str(x)) else "") + "------"
            
            elif column_name == "PRODUCTGROUPID":
                df_transfer[column_name] = (df_input["PRODUCTGROUPID"].astype(str)).apply(lambda x: re.search(r'\d+', str(x)).group() if re.search(r'\d+', str(x)) else "")



            elif column_name == "LANGUAGEID":
                A = "EN-US"
                df_transfer[column_name] = A
            elif column_name == "DISPLAYORDER":
                B = 0
                df_transfer[column_name] = B
            elif column_name == "ISUNITCOSTAUTOMATICALLYUPDATED":
                C = "YES"
                df_transfer[column_name] = C

            elif column_name == "PRODUCTVARIANTNUMBER":
                df_transfer[column_name] = (
                        df_input["ITEMNUMBER"].astype(str) + ": :" +
                        df_input["PRODUCTCOLORID"].astype(str) + ":" +
                        df_input["PRODUCTSIZEID"].astype(str) + ":" +
                        df_input["PRODUCTSTYLEID"].astype(str) + ":"
                )

            elif column_name == "PRODUCTLIFECYCLESTATEID":
                D = ""
                df_transfer[column_name] = D
            elif column_name == "PRODUCTQUANTITY":
                E = ""
                df_transfer[column_name] = E
            elif column_name == "PRODUCTQUANTITYUNITSYMBOL":
                E = "PCS"
                df_transfer[column_name] = E
            
            # Check if "BOM" is available in "ITEMMODELGROUPID"
            elif column_name == "BOMUNITSYMBOL":
                if df_input["ITEMMODELGROUPID"].eq("BOM").any():
                    bom_mask = df_input["ITEMMODELGROUPID"] == "BOM"

                    # Transfer values for "BOMUNITSYMBOL" where "ITEMMODELGROUPID" is "BOM"
                    df_transfer.loc[bom_mask, "BOMUNITSYMBOL"] = df_input.loc[bom_mask, "INVENTORYUNITSYMBOL"]

                    # Transfer values for "CALCULATIONGROUPID" based on "PRODUCTTYPE" where "ITEMMODELGROUPID" is "BOM"
                    df_transfer.loc[bom_mask, "COSTCALCULATIONGROUPID"] = df_input.loc[bom_mask, "PRODUCTTYPE"].apply(lambda x: "PCG" if x == "ITEM" else "SCG" if x == "SRV" else "")
                else:
                    # If "BOM" is not available in "ITEMMODELGROUPID," leave both columns empty
                    pass

            



            
         # Define output_path before the print statement
        output_path = f"{output_folder}/{output_file}"
        
        # Prompt user to remove duplicates
        remove_duplicates = messagebox.askquestion(f"Remove Duplicates",
                                                   f"Do you want to remove duplicates for file '{output_file}'?")
        if remove_duplicates == "yes":
            df_transfer.drop_duplicates(inplace=True)
            print(f"Data transferred, duplicates removed, and saved to {output_path}")
        else:
            print(f"Data transferred and saved to {output_path}")

        df_transfer.to_excel(output_path, index=False)

def browse_input_file():
    input_file = filedialog.askopenfilename(title="Select Input Excel File", filetypes=[("Excel Files", "*.xlsx")])
    input_file_entry.delete(0, tk.END)
    input_file_entry.insert(0, input_file)

def browse_output_folder():
    output_folder = filedialog.askdirectory(title="Select Output Folder")
    output_folder_entry.delete(0, tk.END)
    output_folder_entry.insert(0, output_folder)

app = tk.Tk()
app.title("Excel Data Transfer")

# Input File Entry
input_file_label = Label(app, text="Input Excel File:")
input_file_label.pack()
input_file_entry = tk.Entry(app)
input_file_entry.pack()

input_file_button = tk.Button(app, text="Browse", command=browse_input_file)
input_file_button.pack()

# Checkboxes for selecting output files
selected_files = [IntVar() for _ in range(len(output_files))]

output_files_label = Label(app, text="Select Output Files:")
output_files_label.pack()

for i, file_name in enumerate(output_files):
    file_checkbox = Checkbutton(app, text=file_name, variable=selected_files[i], onvalue=1, offvalue=0)
    file_checkbox.select()  # Select all files by default
    file_checkbox.pack()

# Output Folder Entry
output_folder_label = Label(app, text="Output Folder:")
output_folder_label.pack()
output_folder_entry = tk.Entry(app)
output_folder_entry.pack()

output_folder_button = tk.Button(app, text="Browse", command=browse_output_folder)
output_folder_button.pack()

# Radio buttons for BARCODESETUPID
default_barcodesetupid = IntVar()
default_barcodesetupid.set(1)  # Default value
barcode_label = Label(app, text="Select default value for BARCODESETUPID:")
barcode_label.pack()

barcode_option1 = tk.Radiobutton(app, text="JDOT", variable=default_barcodesetupid, value=1)
barcode_option1.pack()

barcode_option2 = tk.Radiobutton(app, text="ALM", variable=default_barcodesetupid, value=2)
barcode_option2.pack()

barcode_option3 = tk.Radiobutton(app, text="UI", variable=default_barcodesetupid, value=3)
barcode_option3.pack()

# Radio buttons for PRODUCTCATEGORYHIERARCHYNAME
default_productcategoryname = IntVar()
default_productcategoryname.set(1)  # Default value
productcategoryname_label = Label(app, text="Select default value for PRODUCTCATEGORYHIERARCHYNAME:")
productcategoryname_label.pack()

productcategoryname_option1 = tk.Radiobutton(app, text="RETAIL CATEGORY", variable=default_productcategoryname, value=1)
productcategoryname_option1.pack()

productcategoryname_option2 = tk.Radiobutton(app, text="PROCUREMENT CATEGORY", variable=default_productcategoryname, value=2)
productcategoryname_option2.pack()

# Checkbox for custom transformation
apply_custom_transformation = IntVar()
apply_custom_transformation_checkbox = tk.Checkbutton(app, text="Apply Transformation For % values", variable=apply_custom_transformation)
apply_custom_transformation_checkbox.pack()


transfer_button = tk.Button(app, text="Transfer Data", command=lambda: transfer_data(input_file_entry.get(),
                                                                                    output_folder_entry.get(),
                                                                                    default_barcodesetupid,
                                                                                    default_productcategoryname,
                                                                                    apply_custom_transformation,
                                                                                    selected_files))
transfer_button.pack()

app.mainloop()


# In[ ]:




