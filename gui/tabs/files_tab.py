import tkinter as tk
from tkinter import ttk
from ..gui_utils import create_labeled_entry

class FilesTab:
    def __init__(self, parent, gui):
        self.parent = parent
        self.gui = gui
        self.setup()

    def setup(self):
        files_frame = ttk.LabelFrame(self.parent, text="About me")
        files_frame.pack(fill='x', padx=10, pady=10)

        ttk.Label(files_frame, text="Email: sixes2010@gmail.com").grid(row=3, column=0, sticky='w', padx=5, pady=2)
        ttk.Label(files_frame, text="Phone: 185-666-81820").grid(row=4, column=0, sticky='w', padx=5, pady=2)
        #create_labeled_entry(files_frame, "Email: sixes2010@gmail.com", self.gui.excel_file_path, 0, 0, width=50)
        #create_labeled_entry(files_frame, "Phone: 185-666-81820", self.gui.output_file_path, 1, 0, width=50)
        #create_labeled_entry(files_frame, "Output Aux Document:", self.gui.output_aux_file_path, 2, 0, width=50)