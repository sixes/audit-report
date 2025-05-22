import tkinter as tk
from tkinter import ttk
from ..gui_utils import create_labeled_entry

class CompanyTab:
    def __init__(self, parent, gui):
        self.parent = parent
        self.gui = gui
        self.setup()

    def setup(self):
        print(f'CompanyTab setup')

        # Configure global style for disabled Entry widgets (fallback)
        style = ttk.Style()
        style.map('Disabled.TEntry',
                  background=[('disabled', '#d9d9d9')],
                  fieldbackground=[('disabled', '#d9d9d9')],
                  foreground=[('disabled', '#a3a3a3')])
        style.configure('Disabled.TEntry',
                       background='#d9d9d9',
                       fieldbackground='#d9d9d9',
                       foreground='#a3a3a3')

        # Configure custom style for disabled entries
        style.configure('CustomDisabled.TEntry',
                       background='#d9d9d9',
                       fieldbackground='#d9d9d9',
                       foreground='#a3a3a3')
        style.map('CustomDisabled.TEntry',
                  background=[('disabled', '#d9d9d9')],
                  fieldbackground=[('disabled', '#d9d9d9')],
                  foreground=[('disabled', '#a3a3a3')])

        # Configure custom style for enabled entries
        style.configure('CustomEnabled.TEntry',
                       background='#ffffff',
                       fieldbackground='#ffffff',
                       foreground='#000000')
        style.map('CustomEnabled.TEntry',
                  background=[('active', '#ffffff'), ('!disabled', '#ffffff')],
                  fieldbackground=[('active', '#ffffff'), ('!disabled', '#ffffff')],
                  foreground=[('active', '#000000'), ('!disabled', '#000000')])

        company_info_frame = ttk.LabelFrame(self.parent, text="Company Information")
        company_info_frame.pack(fill='x', padx=10, pady=5)

        create_labeled_entry(company_info_frame, "Company Name (English):", self.gui.company_name_en, 0, 0, width=80)
        create_labeled_entry(company_info_frame, "Company Name (Chinese):", self.gui.company_name_cn, 1, 0, width=80)
        create_labeled_entry(company_info_frame, "Company Address:", self.gui.company_address, 2, 0, width=80)
        create_labeled_entry(company_info_frame, "Business Description:", self.gui.business_description, 3, 0, width=80)
        create_labeled_entry(company_info_frame, "Additional Business Description:", self.gui.additional_business_description, 4, 0, width=80)
        create_labeled_entry(company_info_frame, "BR No:", self.gui.br_no, 5, 0)
        create_labeled_entry(company_info_frame, "Last Day of Year:", self.gui.last_day_of_year, 6, 0)
        create_labeled_entry(company_info_frame, "Date of Incorporation:", self.gui.date_of_incorporation, 7, 0)

        options_frame = ttk.LabelFrame(self.parent, text="Company Options")
        options_frame.pack(fill='x', padx=10, pady=5)

        ttk.Checkbutton(options_frame, text="Company has changed name", variable=self.gui.has_name_changed,
                        command=self.gui.toggle_name_change_fields).grid(row=0, column=0, columnspan=2, sticky='w', padx=5, pady=2)

        self.gui.old_company_name_entry = create_labeled_entry(options_frame, "Old Company Name:", self.gui.old_company_name, 1, 0, width=20, state='disabled')
        self.gui.passed_date_entry = create_labeled_entry(options_frame, "Passed Date:", self.gui.passed_date, 3, 0, width=20, state='disabled')
        self.gui.new_company_name_entry = create_labeled_entry(options_frame, "New Company Name:", self.gui.new_company_name, 1, 2, width=20, state='disabled')
        self.gui.effective_date_entry = create_labeled_entry(options_frame, "Effective Date:", self.gui.effective_date, 3, 2, width=20, state='disabled')

        # Store name change entries and apply initial disabled style
        self.name_change_entries = [
            self.gui.old_company_name_entry,
            self.gui.passed_date_entry,
            self.gui.new_company_name_entry,
            self.gui.effective_date_entry
        ]
        for entry in self.name_change_entries:
            if isinstance(entry, ttk.Entry):
                entry.configure(style='CustomDisabled.TEntry')
            elif hasattr(entry, 'entry') and isinstance(entry.entry, ttk.Entry):
                entry.entry.configure(style='CustomDisabled.TEntry')

        # Trace to update name change entry styles
        def update_name_change_styles(*args):
            style_name = 'CustomEnabled.TEntry' if self.gui.has_name_changed.get() else 'CustomDisabled.TEntry'
            state = 'normal' if self.gui.has_name_changed.get() else 'disabled'
            for entry in self.name_change_entries:
                if isinstance(entry, ttk.Entry):
                    entry.configure(style=style_name, state=state)
                elif hasattr(entry, 'entry') and isinstance(entry.entry, ttk.Entry):
                    entry.entry.configure(style=style_name, state=state)

        self.gui.has_name_changed.trace_add('write', update_name_change_styles)

        ttk.Label(options_frame, text="Taxation:").grid(row=5, column=0, sticky='w', padx=5, pady=2)
        tax_frame = ttk.Frame(options_frame)
        tax_frame.grid(row=5, column=1, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(tax_frame, text="亏损不用交税", value="1", variable=self.gui.tax_opt).pack(side='left', padx=2)
        ttk.Radiobutton(tax_frame, text="弥补亏损不用交税", value="2", variable=self.gui.tax_opt).pack(side='left', padx=2)
        ttk.Radiobutton(tax_frame, text="盈利交税", value="3", variable=self.gui.tax_opt).pack(side='left', padx=2)

        ttk.Label(options_frame, text="Capital Change:").grid(row=6, column=0, sticky='w', padx=5, pady=2)
        capital_frame = ttk.Frame(options_frame)
        capital_frame.grid(row=6, column=1, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(capital_frame, text="No Change", value="no_change", variable=self.gui.capital_increase).pack(side='left', padx=2)
        ttk.Radiobutton(capital_frame, text="Increase", value="increase", variable=self.gui.capital_increase).pack(side='left', padx=2)
        ttk.Radiobutton(capital_frame, text="Decrease", value="decrease", variable=self.gui.capital_increase).pack(side='left', padx=2)

        # Add shares numbers entries on the same row, below Capital Change
        shares_curr_entry = create_labeled_entry(options_frame, "Shares (Current Year):", self.gui.shares_curr, 7, 0, width=15)
        self.gui.validate_integer(self.gui.shares_curr, shares_curr_entry)
        shares_prev_entry = create_labeled_entry(options_frame, "Shares (Previous Year):", self.gui.shares_prev, 7, 2, width=15)
        self.gui.validate_integer(self.gui.shares_prev, shares_prev_entry)

        ttk.Checkbutton(options_frame, text="Has Related Party", variable=self.gui.has_related_party).grid(row=8, column=0, sticky='w', padx=5, pady=2)
        ttk.Checkbutton(options_frame, text="Investment in Company", variable=self.gui.investment_in_company,
                        command=lambda: self.gui.enforce_investment_exclusivity('company')).grid(row=8, column=2, sticky='w', padx=5, pady=2)
        ttk.Checkbutton(options_frame, text="Investment in Security", variable=self.gui.investment_in_security,
                        command=lambda: self.gui.enforce_investment_exclusivity('security')).grid(row=8, column=3, sticky='w', padx=5, pady=2)

        self.gui.ultimate_company_frame = ttk.LabelFrame(self.parent, text="Ultimate Company Details")
        self.gui.ultimate_company_frame.pack(fill='x', padx=10, pady=5)

        # Create a frame to hold the checkbox and radiobuttons in the same row
        ultimate_top_frame = ttk.Frame(self.gui.ultimate_company_frame)
        ultimate_top_frame.pack(fill='x', padx=5, pady=5)

        # Place the "Has Ultimate Company" checkbox on the left
        ttk.Checkbutton(ultimate_top_frame, text="Has Ultimate Company:", variable=self.gui.has_ultimate_company,
                        command=self.gui.toggle_ultimate_company_fields).pack(side='left', padx=5)

        # Frame for radiobuttons
        self.gui.ultimate_option_frame = ttk.Frame(ultimate_top_frame)
        self.gui.ultimate_option_frame.pack(side='left', padx=5)

        ttk.Radiobutton(self.gui.ultimate_option_frame, text="One Company, One Location", value="option1",
                        variable=self.gui.ultimate_company_option, command=self.gui.update_ultimate_company_fields).pack(side='left', padx=5)
        ttk.Radiobutton(self.gui.ultimate_option_frame, text="Two Companies, Two Locations", value="option2",
                        variable=self.gui.ultimate_company_option, command=self.gui.update_ultimate_company_fields).pack(side='left', padx=5)
        ttk.Radiobutton(self.gui.ultimate_option_frame, text="Two Companies, One Location", value="option3",
                        variable=self.gui.ultimate_company_option, command=self.gui.update_ultimate_company_fields).pack(side='left', padx=5)

        self.gui.ultimate_details_frame = ttk.Frame(self.gui.ultimate_company_frame)
        self.gui.ultimate_details_frame.pack(fill='x', padx=5, pady=5)

        self.gui.ultimate_company_name1_entry = create_labeled_entry(self.gui.ultimate_details_frame, "Company Name 1:", self.gui.ultimate_company_name1, 0, 0, width=30)
        self.gui.ultimate_company_location1_entry = create_labeled_entry(self.gui.ultimate_details_frame, "Location 1:", self.gui.ultimate_company_location1, 0, 3, width=30)
        self.gui.ultimate_company_name2_entry = create_labeled_entry(self.gui.ultimate_details_frame, "Company Name 2:", self.gui.ultimate_company_name2, 1, 0, width=30)
        self.gui.ultimate_company_location2_entry = create_labeled_entry(self.gui.ultimate_details_frame, "Location 2:", self.gui.ultimate_company_location2, 1, 3, width=30)

        # Store ultimate company entries
        self.ultimate_company_entries = [
            self.gui.ultimate_company_name1_entry,
            self.gui.ultimate_company_location1_entry,
            self.gui.ultimate_company_name2_entry,
            self.gui.ultimate_company_location2_entry
        ]

        # Apply initial disabled style
        for entry in self.ultimate_company_entries:
            if isinstance(entry, ttk.Entry):
                entry.configure(style='CustomDisabled.TEntry', state='disabled')
            elif hasattr(entry, 'entry') and isinstance(entry.entry, ttk.Entry):
                entry.entry.configure(style='CustomDisabled.TEntry', state='disabled')

        # Trace to update ultimate company entry states and styles
        def update_ultimate_company_entries(*args):
            if not self.gui.has_ultimate_company.get():
                # All entries disabled if checkbox is unchecked
                for entry in self.ultimate_company_entries:
                    if isinstance(entry, ttk.Entry):
                        entry.configure(state='disabled', style='CustomDisabled.TEntry')
                    elif hasattr(entry, 'entry') and isinstance(entry.entry, ttk.Entry):
                        entry.entry.configure(state='disabled', style='CustomDisabled.TEntry')
            else:
                # Enable/disable based on selected option
                option = self.gui.ultimate_company_option.get()
                configs = {
                    'option1': [('normal', 'CustomEnabled.TEntry'), ('normal', 'CustomEnabled.TEntry'), 
                                ('disabled', 'CustomDisabled.TEntry'), ('disabled', 'CustomDisabled.TEntry')],
                    'option2': [('normal', 'CustomEnabled.TEntry'), ('normal', 'CustomEnabled.TEntry'), 
                                ('normal', 'CustomEnabled.TEntry'), ('normal', 'CustomEnabled.TEntry')],
                    'option3': [('normal', 'CustomEnabled.TEntry'), ('normal', 'CustomEnabled.TEntry'), 
                                ('normal', 'CustomEnabled.TEntry'), ('disabled', 'CustomDisabled.TEntry')]
                }.get(option, [('disabled', 'CustomDisabled.TEntry')] * 4)  # Fallback if option is invalid
                for entry, (state, style_name) in zip(self.ultimate_company_entries, configs):
                    if isinstance(entry, ttk.Entry):
                        entry.configure(state=state, style=style_name)
                    elif hasattr(entry, 'entry') and isinstance(entry.entry, ttk.Entry):
                        entry.entry.configure(state=state, style=style_name)

        # Add traces for both checkbox and radiobutton variables
        self.gui.has_ultimate_company.trace_add('write', update_ultimate_company_entries)
        self.gui.ultimate_company_option.trace_add('write', update_ultimate_company_entries)

        self.gui.toggle_ultimate_company_fields()

        # Create a parent Frame to hold both LabelFrames in the same row
        row_frame = ttk.Frame(self.parent)
        row_frame.pack(fill='x', padx=10, pady=5)

        # First LabelFrame (Directors)
        directors_frame = ttk.LabelFrame(row_frame, text="Directors (One per line)")
        directors_frame.pack(side='left', fill='x', expand=True, padx=5)

        self.gui.directors_text = tk.Text(directors_frame, height=5, width=40)
        self.gui.directors_text.pack(fill='x', padx=5, pady=5)
        self.gui.directors_text.insert('1.0', self.gui.directors.get())
        self.gui.directors_text.bind("<KeyRelease>", self.gui.update_directors)

        # Second LabelFrame (Shareholders)
        shareholders_frame = ttk.LabelFrame(row_frame, text="Shareholders (One per line)")
        shareholders_frame.pack(side='left', fill='x', expand=True, padx=5)

        self.gui.shareholders_text = tk.Text(shareholders_frame, height=5, width=40)
        self.gui.shareholders_text.pack(fill='x', padx=5, pady=5)
        self.gui.shareholders_text.insert('1.0', self.gui.shareholders.get())
        self.gui.shareholders_text.bind("<KeyRelease>", self.gui.update_shareholders)