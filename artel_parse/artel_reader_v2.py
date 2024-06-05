import pandas as pd
import numpy as np
import ipywidgets as ipw
from IPython.display import display
from datetime import datetime
import xml.etree.ElementTree as ET
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory


class XMLFormatter:
    def __init__(self):
        # Initialize date format for naming output files
        self.now = datetime.today().strftime("%y%m%d")
        
        # Output widget for displaying data
        self.show_df = ipw.Output()
        
        # DataFrame to hold all groups' data
        self.all_groups_df = pd.DataFrame()
        
        # Hide Tkinter main window
        self.window = Tk()
        self.window.withdraw()
        
        # Prompt user to select the XML file
        self.file_path = askopenfilename(title="Choose xml file")
        
        # Prompt user to select the directory to save the output file
        self.drop_path = askdirectory(title="Where should the file go?")
        
        # Parse the selected XML file
        self.tree = ET.parse(self.file_path)
        self.root = self.tree.getroot()
        
        # Format the data from the XML file
        self.format_data()
        
        # Destroy the Tkinter window after processing
        self.window.destroy()

    def format_data(self):
        # Get list of top-level tags in the XML file
        id_list = [tag.tag for tag in self.root]
        
        # Determine the run ID based on XML structure
        if "Group" in id_list:
            run_id = "Group"
        elif "Plate" in id_list:
            run_id = "Plate"
        else:
            run_id = ""
        
        # Show an error if run ID is not found
        if run_id == "":
            messagebox.showerror(title="Warning...", message="Something may have changed with the XML file.")
        else:
            # Iterate over each group (or plate) in the XML file
            for group in self.root.iter(run_id):
                # Extract column names
                col_list = [col_num.text for column in group.iter("Columns") for col_num in column.iter('Name')]
                col_len = len(col_list)
                
                # Extract unique row names
                row_list = sorted(list(set([row.text for row in group.iter("Row")])))
                row_len = len(row_list)
                
                # Extract values and reshape into a matrix
                value_list = np.array([value.text for value in group.iter("Value")])
                value_matrix = value_list.reshape(row_len, col_len)
                
                # Create a DataFrame for the current group
                group_df = pd.DataFrame(value_matrix, columns=col_list, index=row_list)
                
                # Concatenate the current group DataFrame with the main DataFrame
                self.all_groups_df = pd.concat([self.all_groups_df, group_df], axis=1)

            # Display the formatted data
            display(self.show_df, ipw.HTML("<h3><b><u>Output display</u></b></h3>"), self.all_groups_df)
            
            # Attempt to save the DataFrame to an Excel file, handle file open errors
            file_open = True
            while file_open:
                try:
                    pd.DataFrame.to_excel(self.all_groups_df, f"{self.drop_path}/Artel_data_{self.now}.xlsx")
                except PermissionError:
                    messagebox.showerror(title="Oops...", message="Do you have the file open?")
                else:
                    messagebox.showinfo(title="Way to go!", message="Data has been formatted and output!")
                    file_open = False


if __name__ == "__main__":
    # Run the XMLFormatter if the script is executed directly
    XMLFormatter()

