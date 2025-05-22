import tkinter as tk
from tkinter import ttk
from ..gui_utils import create_labeled_entry

class GeneralTab:
    def __init__(self, parent, gui):
        self.parent = parent
        self.gui = gui
        self.setup()

    def setup(self):
        audit_info_frame = ttk.LabelFrame(self.parent, text="Audit Information")
        audit_info_frame.pack(fill='x', padx=10, pady=10)
        # Adding Audit Type dropdown
        ttk.Label(audit_info_frame, text="Audit Type:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        audit_type_combo = ttk.Combobox(audit_info_frame, textvariable=self.gui.audit_type, values=["LAI", "LEUNG", "WH", "WOCP", "WU"], state="readonly")
        audit_type_combo.grid(row=1, column=1, sticky='w', padx=5, pady=5)

        ttk.Label(audit_info_frame, text="First Year:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        opinion_frame = ttk.Frame(audit_info_frame)
        opinion_frame.grid(row=2, column=1, sticky='w', padx=5, pady=5)
        ttk.Radiobutton(opinion_frame, text="Yes", value=True, variable=self.gui.first_year).pack(side='left', padx=5)
        ttk.Radiobutton(opinion_frame, text="No", value=False, variable=self.gui.first_year).pack(side='left', padx=5)

        ttk.Label(audit_info_frame, text="Audit Opinion:").grid(row=3, column=0, sticky='w', padx=5, pady=5)
        opinion_frame = ttk.Frame(audit_info_frame)
        opinion_frame.grid(row=3, column=1, sticky='w', padx=5, pady=5)
        ttk.Radiobutton(opinion_frame, text="Opinion", value="Opinion", variable=self.gui.audit_opinion).pack(side='left', padx=5)
        ttk.Radiobutton(opinion_frame, text="Qualified Opinion", value="Qualified Opinion", variable=self.gui.audit_opinion).pack(side='left', padx=5)

        create_labeled_entry(audit_info_frame, "Approval Date:", self.gui.approval_date, 4, 0)

        year_frame = ttk.LabelFrame(self.parent, text="Reporting Years")
        year_frame.pack(fill='x', padx=10, pady=10)
        ttk.Label(year_frame, text="Current Year:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.gui.year_var = tk.StringVar(value=str(self.gui.current_year))
        self.gui.year_var.trace_add("write", self.gui.on_year_change)
        self.gui.year_spinbox = ttk.Spinbox(year_frame, from_=2000, to=2100, width=10, textvariable=self.gui.year_var)
        self.gui.year_spinbox.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        self.gui.validate_integer(self.gui.year_var, self.gui.year_spinbox)

        ttk.Label(year_frame, text="Previous Year:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.gui.prev_year_label = ttk.Label(year_frame, text=str(self.gui.previous_year))
        self.gui.prev_year_label.grid(row=1, column=1, sticky='w', padx=5, pady=5)

        self.gui.year_spinbox.bind("<KeyRelease>", self.gui.update_previous_year)
        self.gui.year_spinbox.bind("<<Increment>>", self.gui.update_previous_year)
        self.gui.year_spinbox.bind("<<Decrement>>", self.gui.update_previous_year)

        business_frame = ttk.LabelFrame(self.parent, text="Business Type")
        business_frame.pack(fill='x', padx=10, pady=10)

        self.gui.business_type.trace_add("write", self.gui.update_business_description)
        ttk.Radiobutton(business_frame, text="General Trading", value="general trading",
                        variable=self.gui.business_type).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        ttk.Radiobutton(business_frame, text="Services", value="services",
                        variable=self.gui.business_type).grid(row=0, column=1, sticky='w', padx=5, pady=5)
        ttk.Radiobutton(business_frame, text="Agency Services", value="agency services",
                        variable=self.gui.business_type).grid(row=0, column=2, sticky='w', padx=5, pady=5)
        ttk.Radiobutton(business_frame, text="Investment Holding", value="investment holding",
                        variable=self.gui.business_type).grid(row=0, column=3, sticky='w', padx=5, pady=5)
        ttk.Radiobutton(business_frame, text="Dormant", value="dormant",
                        variable=self.gui.business_type).grid(row=0, column=4, sticky='w', padx=5, pady=5)

        """
        shares_frame = ttk.LabelFrame(self.parent, text="Share Numbers")
        shares_frame.pack(fill='x', padx=10, pady=10)

        shares_curr_entry = create_labeled_entry(shares_frame, "Shares (Current Year):", self.gui.shares_curr, 0, 0, width=15)
        self.gui.validate_integer(self.gui.shares_curr, shares_curr_entry)

        shares_prev_entry = create_labeled_entry(shares_frame, "Shares (Previous Year):", self.gui.shares_prev, 1, 0, width=15)
        self.gui.validate_integer(self.gui.shares_prev, shares_prev_entry)
        """
        currency_frame = ttk.LabelFrame(self.parent, text="Currency")
        currency_frame.pack(fill='x', padx=10, pady=10)

        self.gui.currency_choice = tk.StringVar(value="HKD")
        self.gui.currency_choice.trace_add("write", self.gui.update_currency)

        ttk.Label(currency_frame, text="Currency:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.gui.currency_combobox = ttk.Combobox(currency_frame, textvariable=self.gui.currency_choice, state='readonly')
        self.gui.currency_combobox['values'] = ['HKD', 'USD', 'RMB']
        self.gui.currency_combobox.grid(row=0, column=1, sticky='w', padx=5, pady=5)

        inventory_frame = ttk.LabelFrame(self.parent, text="Inventory Valuation Method")
        inventory_frame.pack(fill='x', padx=10, pady=10)

        ttk.Label(inventory_frame, text="Valuation Method:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.gui.inventory_combobox = ttk.Combobox(inventory_frame, textvariable=self.gui.inventory_valuation, state='readonly')
        self.gui.inventory_combobox['values'] = ['FIFO', 'Weighted Average', 'Specific Identification']
        self.gui.inventory_combobox.grid(row=0, column=1, sticky='w', padx=5, pady=5)

        files_frame = ttk.LabelFrame(self.parent, text="File Selection")
        files_frame.pack(fill='x', padx=10, pady=10)

        create_labeled_entry(files_frame, "Trial Balance Excel File:", self.gui.excel_file_path, 0, 0, width=50)
        ttk.Button(files_frame, text="Browse...", command=self.gui.browse_excel).grid(row=0, column=2, sticky='w', padx=5, pady=5)

        create_labeled_entry(files_frame, "Output Word Document:", self.gui.output_file_path, 1, 0, width=50)
        ttk.Button(files_frame, text="Browse...", command=self.gui.browse_output).grid(row=1, column=2, sticky='w', padx=5, pady=5)

        create_labeled_entry(files_frame, "Output Aux Document:", self.gui.output_aux_file_path, 2, 0, width=50)
        ttk.Button(files_frame, text="Browse...", command=self.gui.browse_output_aux).grid(row=2, column=2, sticky='w', padx=5, pady=5)
        
        manage_categories_frame = ttk.LabelFrame(self.parent, text="Advanced Settings")
        manage_categories_frame.pack(fill='x', padx=10, pady=10)

        ttk.Button(manage_categories_frame, text="Manage Categories", command=self.gui.open_manage_categories_dialog).pack(pady=5)