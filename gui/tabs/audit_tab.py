import tkinter as tk
from tkinter import ttk
from ..gui_utils import create_labeled_entry

class AuditTab:
    def __init__(self, parent, gui):
        self.parent = parent
        self.gui = gui
        self.setup()

    def setup(self):
        audit_info_frame = ttk.LabelFrame(self.parent, text="Audit Information")
        audit_info_frame.pack(fill='x', padx=10, pady=10)

        """
        create_labeled_entry(audit_info_frame, "Audit Firm:", self.gui.audit_firm, 0, 0)
        create_labeled_entry(audit_info_frame, "Auditor Name:", self.gui.auditor_name, 2, 0)
        create_labeled_entry(audit_info_frame, "Auditor License No:", self.gui.auditor_license, 3, 0)
        """

        # Adding Audit Type dropdown
        ttk.Label(audit_info_frame, text="Audit Type:").grid(row=4, column=0, sticky='w', padx=5, pady=5)
        audit_type_combo = ttk.Combobox(audit_info_frame, textvariable=self.gui.audit_type, values=["WOCP", "WH", "LAI", "LEUNG", "WU"], state="readonly")
        audit_type_combo.grid(row=4, column=1, sticky='w', padx=5, pady=5)

        ttk.Label(audit_info_frame, text="Audit Opinion:").grid(row=5, column=0, sticky='w', padx=5, pady=5)
        opinion_frame = ttk.Frame(audit_info_frame)
        opinion_frame.grid(row=5, column=1, sticky='w', padx=5, pady=5)
        ttk.Radiobutton(opinion_frame, text="Opinion", value="Opinion", variable=self.gui.audit_opinion).pack(side='left', padx=5)
        ttk.Radiobutton(opinion_frame, text="Qualified Opinion", value="Qualified Opinion", variable=self.gui.audit_opinion).pack(side='left', padx=5)