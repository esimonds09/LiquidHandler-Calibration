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
        self.now = datetime.today().strftime("%y%m%d")
        self.show_df = ipw.Output()
        self.all_groups_df = pd.DataFrame()
        self.window = Tk()
        self.window.withdraw()
        self.file_path = askopenfilename(title="Choose xml file")
        self.drop_path = askdirectory(title="Where should the file go?")
        self.tree = ET.parse(self.file_path)
        self.root = self.tree.getroot()
        self.format_data()
        self.window.destroy()

    def format_data(self):
        id_list = [tag.tag for tag in self.root]
        if "Group" in id_list:
            run_id = "Group"
        elif "Plate" in id_list:
            run_id = "Plate"
        else:
            run_id = ""
        if run_id == "":
            messagebox.showerror(title="Warning...", message="Something may have changed with the XML file.")
        else:
            for group in self.root.iter(run_id):
                col_list = [col_num.text for column in group.iter("Columns") for col_num in column.iter('Name')]
                col_len = len(col_list)
                row_list = sorted(list(set([row.text for row in group.iter("Row")])))
                row_len = len(row_list)
                value_list = np.array([value.text for value in group.iter("Value")])
                value_matrix = value_list.reshape(row_len, col_len)
                group_df = pd.DataFrame(value_matrix, columns=col_list, index=row_list)
                self.all_groups_df = pd.concat([self.all_groups_df, group_df], axis=1)

            display(self.show_df, ipw.HTML("<h3><b><u>Output display</u></b></h3>"), self.all_groups_df)
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
    XMLFormatter()
