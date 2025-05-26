import os
import sys
import time
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime
# import pandas as pd  # Moved to generate_aux_report and generate_report
# import xml.sax.saxutils as saxutils  # Uncomment if needed for escaping

# Setup file-based logging
start_time = time.time()
log_file = os.path.join(os.path.dirname(sys.executable), "startup_log.txt") if hasattr(sys, 'frozen') else "startup_log.txt"
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logging.info(f"Start imports: {start_time:.3f} seconds")

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from document_generator import DocumentGenerator
from exceptions import *
from .category_manager import CategoryManager
from .tabs.general_tab import GeneralTab
from .tabs.company_tab import CompanyTab
# from .tabs.audit_tab import AuditTab
# from .tabs.files_tab import FilesTab

logging.info(f"Finished imports: {time.time() - start_time:.3f} seconds")

class AuditReportGUI(tk.Tk):
    def __init__(self):
        init_start = time.time()
        logging.info(f"Start AuditReportGUI.__init__: {init_start - start_time:.3f} seconds")

        super().__init__()
        self.title("深圳好景商务公司 * Audit Report Generator")
        self.geometry("1024x900")

        logging.info(f"Window setup: {time.time() - init_start:.3f} seconds")

        # Default values
        self.current_year = datetime.datetime.now().year - 1
        self.previous_year = self.current_year - 1

        # Variables
        self.company_name_en = tk.StringVar(value="")
        self.company_name_cn = tk.StringVar(value="")
        self.company_address = tk.StringVar(value="")
        self.business_description = tk.StringVar(value="general trading")
        self.additional_business_description = tk.StringVar(value="")
        self.br_no = tk.StringVar(value="")
        self.last_day_of_year = tk.StringVar(value=f"31 December {self.current_year}")
        self.date_of_incorporation = tk.StringVar(value="")
        self.audit_firm = tk.StringVar(value="Deloitte")
        self.approval_date = tk.StringVar(value="")
        self.auditor_name = tk.StringVar(value="Auditor Name")
        self.auditor_license = tk.StringVar(value="CPA12345")
        self.currency_desc = tk.StringVar(value="Hong Kong dollars")
        self.currency_full_desc = tk.StringVar(value="Hong Kong dollars (HK$)")
        self.currency = tk.StringVar(value="HK$")
        self.directors = tk.StringVar(value="")
        self.shareholders = tk.StringVar(value="")
        self.business_type = tk.StringVar(value="general trading")
        self.shares_curr = tk.StringVar(value="10000")
        self.shares_prev = tk.StringVar(value="10000")
        self.has_name_changed = tk.BooleanVar(value=False)
        self.passed_date = tk.StringVar()
        self.new_company_name = tk.StringVar()
        self.effective_date = tk.StringVar()
        self.old_company_name = tk.StringVar()
        self.has_related_party = tk.BooleanVar(value=False)
        self.inventory_valuation = tk.StringVar(value="FIFO")
        self.tax_opt = tk.StringVar(value="1")
        self.capital_increase = tk.StringVar(value="no_change")
        self.has_ultimate_company = tk.BooleanVar(value=False)
        self.ultimate_company_option = tk.StringVar(value="option1")
        self.ultimate_company_name1 = tk.StringVar()
        self.ultimate_company_location1 = tk.StringVar()
        self.ultimate_company_name2 = tk.StringVar()
        self.ultimate_company_location2 = tk.StringVar()
        self.investment_in_company = tk.BooleanVar(value=False)
        self.investment_in_security = tk.BooleanVar(value=False)
        self.audit_opinion = tk.StringVar(value="Opinion")
        self.audit_type = tk.StringVar(value="WOCP")
        self.first_year = tk.BooleanVar(value=False)

        self.excel_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar(value=os.path.join(os.path.dirname(os.path.abspath(__file__)), "audit_report_filled.docx"))
        self.output_aux_file_path = tk.StringVar(value=os.path.join(os.path.dirname(os.path.abspath(__file__)), "aux_report_filled.docx"))

        logging.info(f"Variables setup: {time.time() - init_start:.3f} seconds")

        # Optimize notebook style
        style = ttk.Style()
        style.configure("TNotebook", tabfocus=False)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)

        self.general_frame = ttk.Frame(self.notebook)
        self.company_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.general_frame, text="General")
        self.notebook.add(self.company_frame, text="Company Info")

        logging.info(f"Notebook setup: {time.time() - init_start:.3f} seconds")

        GeneralTab(self.general_frame, self)
        CompanyTab(self.company_frame, self)

        logging.info(f"Tab creation: {time.time() - init_start:.3f} seconds")

        """
        # Setup tabs with slight delay to ensure event loop readiness
        def setup_tabs():
            GeneralTab(self.general_frame, self)
            CompanyTab(self.company_frame, self)
            #AuditTab(self.audit_frame, self)
            #FilesTab(self.files_frame, self)
            #AboutContactTab(self.about_frame, self)

            # Pre-render all tabs to cache layouts
            for i in range(self.notebook.index("end")):
                self.notebook.select(i)
                self.update_idletasks()
                # Log widget count for diagnostics
                tab_frame = self.notebook.winfo_children()[i]
                widget_count = self._count_widgets_recursive(tab_frame)
                #print(f"Tab '{self.notebook.tab(i, 'text')}' has {widget_count} widgets")
            self.notebook.select(0)  # Return to first tab

        self.after(0, setup_tabs)

        # Force rendering on tab change to eliminate delay
        def on_tab_change(event):
            start_time = time.time()
            tab_name = event.widget.tab(event.widget.select(), "text")
            # Force immediate rendering
            self.update()
            print(f"Switched to tab '{tab_name}' in {time.time() - start_time:.3f} seconds")
        self.notebook.bind("<<NotebookTabChanged>>", on_tab_change)
        """

        self.buttons_frame = ttk.Frame(self)
        self.buttons_frame.pack(fill='x', padx=10, pady=1)

        self.generate_btn = ttk.Button(self.buttons_frame, text="Generate Audit Report", command=self.generate_report)
        self.generate_btn.pack(side='right', padx=5)

        self.generate_aux_btn = ttk.Button(self.buttons_frame, text="Generate Aux Report", command=self.generate_aux_report)
        self.generate_aux_btn.pack(side='right', padx=5)

        self.status_label = ttk.Label(self, text="Ready")
        self.status_label.pack(side='bottom', fill='x', padx=10, pady=5)

        logging.info(f"Buttons and status setup: {time.time() - init_start:.3f} seconds")

        # Initialize category manager
        self.category_manager = CategoryManager()
        logging.info(f"CategoryManager init: {time.time() - init_start:.3f} seconds")

        self._document_generator = DocumentGenerator(self.category_manager)
        logging.info(f"DocumentGenerator init: {time.time() - init_start:.3f} seconds")

        self.load_categories()
        logging.info(f"load_categories: {time.time() - init_start:.3f} seconds")

        logging.info(f"Finished AuditReportGUI.__init__: {time.time() - init_start:.3f} seconds")

    def _count_widgets_recursive(self, widget):
        """Recursively count all widgets in the widget tree."""
        count = 1  # Count the current widget
        for child in widget.winfo_children():
            count += self._count_widgets_recursive(child)
        return count

    def load_categories(self):
        self.non_current_assets = self.category_manager.categories['non_current_assets']
        self.current_assets = self.category_manager.categories['current_assets']
        self.current_liabilities = self.category_manager.categories['current_liabilities']
        self.non_current_liabilities = self.category_manager.categories['non_current_liabilities']
        self.equity = self.category_manager.categories['equity']
        self.revenue_items = self.category_manager.categories['revenue_items']
        self.cost_of_sales_items = self.category_manager.categories['cost_of_sales_items']
        self.closing_inventories = self.category_manager.categories['closing_inventories']
        self.other_income_items = self.category_manager.categories['other_income_items']
        self.general_admin_expenses_items = self.category_manager.categories['general_admin_expenses_items']
        self.finance_costs_items = self.category_manager.categories['finance_costs_items']
        self.tax_items = self.category_manager.categories['tax_items']

    def toggle_name_change_fields(self):
        state = 'normal' if self.has_name_changed.get() else 'disabled'
        self.passed_date_entry.config(state=state)
        self.new_company_name_entry.config(state=state)
        self.effective_date_entry.config(state=state)
        self.old_company_name_entry.config(state=state)

    def toggle_ultimate_company_fields(self):
        state = 'normal' if self.has_ultimate_company.get() else 'disabled'
        for widget in self.ultimate_option_frame.winfo_children():
            if isinstance(widget, ttk.Radiobutton):
                widget.config(state=state)
        self.update_ultimate_company_fields()

    def update_ultimate_company_fields(self, *args):
        option = self.ultimate_company_option.get()
        base_state = 'normal' if self.has_ultimate_company.get() else 'disabled'
        if option == "option1":
            self.ultimate_company_name1_entry.config(state=base_state)
            self.ultimate_company_location1_entry.config(state=base_state)
            self.ultimate_company_name2_entry.config(state='disabled')
            self.ultimate_company_location2_entry.config(state='disabled')
        elif option == "option2":
            self.ultimate_company_name1_entry.config(state=base_state)
            self.ultimate_company_location1_entry.config(state=base_state)
            self.ultimate_company_name2_entry.config(state=base_state)
            self.ultimate_company_location2_entry.config(state=base_state)
        elif option == "option3":
            self.ultimate_company_name1_entry.config(state=base_state)
            self.ultimate_company_location1_entry.config(state=base_state)
            self.ultimate_company_name2_entry.config(state=base_state)
            self.ultimate_company_location2_entry.config(state='disabled')

    def enforce_investment_exclusivity(self, selected_var):
        if selected_var == 'company' and self.investment_in_company.get():
            self.investment_in_security.set(False)
        elif selected_var == 'security' and self.investment_in_security.get():
            self.investment_in_company.set(False)

    def update_business_description(self, *args):
        business_type = self.business_type.get()
        if business_type == "general trading":
            self.business_description.set("general trading")
        elif business_type == "agency services":
            self.business_description.set("provision of agency services")
        elif business_type == "investment holding":
            self.business_description.set("investment holding")
        elif business_type == "services":
            self.business_description.set("provision of services")
        elif business_type == "dormant":
            self.business_description.set("dormant")

    def update_currency(self, *args):
        currency_choice = self.currency_choice.get()
        if currency_choice == "HKD":
            self.currency.set("HK$")
            self.currency_desc.set("Hong Kong dollars")
            self.currency_full_desc.set("Hong Kong dollars (HK$)")
        elif currency_choice == "USD":
            self.currency.set("US$")
            self.currency_desc.set("United States dollars")
            self.currency_full_desc.set("United States dollars (US$)")
        elif currency_choice == "RMB":
            self.currency.set("RMB")
            self.currency_desc.set("Renminbi Yuan")
            self.currency_full_desc.set("Renminbi Yuan(RMB)")

    def on_year_change(self, *args):
        try:
            current_year = int(self.year_var.get())
            self.previous_year = current_year - 1
            self.prev_year_label.config(text=str(self.previous_year))
            self.last_day_of_year.set(f"31 December {current_year}")
        except ValueError:
            pass

    def update_previous_year(self, event=None):
        try:
            if event and hasattr(event, 'widget'):
                if event.type == '<<Increment>>':
                    current_year = int(self.year_spinbox.get()) + 1
                    self.year_spinbox.set(current_year)
                elif event.type == '<<Decrement>>':
                    current_year = int(self.year_spinbox.get()) - 1
                    self.year_spinbox.set(current_year)
                else:
                    current_year = int(self.year_spinbox.get())
            else:
                current_year = int(self.year_spinbox.get())

            self.previous_year = current_year - 1
            self.prev_year_label.config(text=str(self.previous_year))
            self.last_day_of_year.set(f"31 December {current_year}")
        except ValueError:
            pass

    def update_directors(self, event=None):
        self.directors.set(self.directors_text.get('1.0', 'end-1c'))

    def update_shareholders(self, event=None):
        self.shareholders.set(self.shareholders_text.get('1.0', 'end-1c'))

    def browse_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Trial Balance Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_file_path.set(file_path)

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Audit Report As",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")]
        )
        if file_path:
            self.output_file_path.set(file_path)

    def browse_output_aux(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Aux Report As",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")]
        )
        if file_path:
            self.output_aux_file_path.set(file_path)

    def validate_integer(self, var, entry):
        """Validate that a StringVar contains a valid integer."""
        def callback(*args):
            try:
                value = var.get()
                if value:
                    int(value)
                    entry.config(foreground='black')
                else:
                    entry.config(foreground='black')  # Allow empty
            except ValueError:
                entry.config(foreground='red')
        var.trace_add('write', callback)

    def show_error(self, message, status_text=None):
        """Show an error message and update status label."""
        messagebox.showerror("Error", message)
        self.status_label.config(text=status_text or f"Error: {message}")

    def open_manage_categories_dialog(self):
        from .manage_categories_dialog import ManageCategoriesDialog
        ManageCategoriesDialog(self)

    def generate_aux_report(self):
        if not self.excel_file_path.get():
            self.show_error("Please select a Trial Balance Excel file", "Error: No Excel file selected")
            return

        if not self.output_aux_file_path.get():
            self.show_error("Please specify an output aux file path", "Error: No output aux path specified")
            return

        # Validate shareholders
        shareholders = self.shareholders.get().splitlines()
        shareholders_list = [s.strip() for s in shareholders if s.strip()]
        if not shareholders_list:
            self.show_error("Shareholders list is empty after cleaning")
            return

        try:
            self.current_year = int(self.year_var.get())
        except ValueError:
            self.show_error("Current year must be a valid integer.", "Error: Invalid current year")
            return

        import pandas as pd
        excel = pd.ExcelFile(self.excel_file_path.get())
        sheet_names = excel.sheet_names
        self.previous_year = self.current_year - 1
        self.current_sheet = f"{self.current_year}TB"
        self.previous_sheet = f"{self.previous_year}TB"

        missing_sheets = []
        if self.current_sheet not in sheet_names:
            missing_sheets.append(self.current_sheet)
        if not self.first_year.get() and self.previous_sheet not in sheet_names:
            missing_sheets.append(self.previous_sheet)

        if missing_sheets:
            error_msg = f"The Excel file does not contain the required sheet(s): {', '.join(missing_sheets)}.\n"
            error_msg += f"The Excel file contains the following sheets: {', '.join(sheet_names)}."
            messagebox.showerror("Sheet Not Found", error_msg)
            raise ValueError(f"Required sheets not found: {error_msg}")

        self.status_label.config(text="Generating report... Please wait.")
        self.update()

        output_aux_file_path = self.output_aux_file_path.get()
        print(f"Output aux path from GUI: {output_aux_file_path}")

        try:
            result, error_message = self._document_generator.generate_aux_document(
                last_day_of_year=self.last_day_of_year.get(),
                date_of_incorporation=self.date_of_incorporation.get(),
                first_year=self.first_year.get(),
                business_type=self.business_type.get(),
                aux_output_path=output_aux_file_path,
                company_name_en=self.company_name_en.get(),
                company_name_cn=self.company_name_cn.get(),
                directors=self.directors.get().splitlines(),
                shareholders=shareholders_list,
                currency=self.currency.get(),
                has_stocking_letter=False,
                br_no=self.br_no.get(),
                excel_file=self.excel_file_path.get(),
                current_year=self.current_year,
                audit_type=self.audit_type.get(),
            )

            if result is None:
                self.status_label.config(text=error_message)
                messagebox.showwarning("Warning", error_message)
                return

            if not os.path.exists(output_aux_file_path):
                self.show_error(f"Generated file not found at {output_aux_file_path}", "Error: Generated file not found")
                return

            self.status_label.config(text=f"Report generated successfully: {output_aux_file_path}")

            if messagebox.askyesno("Success", f"Report generated successfully at {output_aux_file_path}. Would you like to open it now?"):
                if sys.platform == 'darwin':
                    os.system(f"open '{output_aux_file_path}'")
                elif sys.platform == 'win32':
                    os.system(f'start "" "{output_aux_file_path}"')
                else:
                    os.system(f"xdg-open '{output_aux_file_path}'")
        except Exception as e:
            raise e

    def generate_report(self):
        if not self.excel_file_path.get():
            self.show_error("Please select a Trial Balance Excel file", "Error: No Excel file selected")
            return

        if not self.output_file_path.get():
            self.show_error("Please specify an output file path", "Error: No output path specified")
            return

        try:
            shares_curr_input = self.shares_curr.get()
            shares_prev_input = self.shares_prev.get()

            shares_curr_value = int(shares_curr_input)
            shares_prev_value = int(shares_prev_input)
            if shares_curr_value < 0 or shares_prev_value < 0:
                raise ValueError("Share numbers must be non-negative.")
        except ValueError as e:
            if "must be non-negative" in str(e):
                self.show_error("Share numbers must be non-negative integers.", "Error: Invalid share numbers")
            else:
                self.show_error("Share numbers must be valid integers.", "Error: Invalid share numbers")
            return

        try:
            current_year = int(self.year_var.get())
        except ValueError:
            self.show_error("Current year must be a valid integer.", "Error: Invalid current year")
            return

        self.status_label.config(text="Generating report... Please wait.")
        self.update()

        excel_file_path = self.excel_file_path.get()
        output_path = self.output_file_path.get()

        print(f"Excel file path from GUI: {excel_file_path}")
        print(f"Output path from GUI: {output_path}")
        print(f"Current year: {current_year}, Previous year: {current_year-1}")
        print(f"Expected sheet names: {current_year}TB and {current_year-1}TB")

        if not os.path.exists(excel_file_path):
            self.show_error(f"Trial balance Excel file not found: {excel_file_path}", "Error: Excel file not found")
            return

        from data_loader import DataLoader
        original_init = DataLoader.__init__

        def custom_init(self, excel_file, first_year=False, current_year=None, non_current_assets=None, current_assets=None,
                current_liabilities=None, non_current_liabilities=None, equity=None,
                revenue_items=None, cost_of_sales_items=None, closing_inventories=None,
                other_income_items=None, general_admin_expenses_items=None,
                finance_costs_items=None, tax_items=None):
            gui_year = current_year

            print(f"Custom init with forced year from GUI: {gui_year}")
            self.excel_file = excel_file
            self.current_year = gui_year
            self.previous_year = gui_year - 1
            self.current_sheet = f"{self.current_year}TB"
            self.previous_sheet = f"{self.previous_year}TB"
            print(f"Using sheet names: {self.current_sheet} and {self.previous_sheet}")

            try:
                import pandas as pd
                excel = pd.ExcelFile(excel_file)
                sheet_names = excel.sheet_names

                missing_sheets = []
                if self.current_sheet not in sheet_names:
                    missing_sheets.append(self.current_sheet)
                if not first_year and self.previous_sheet not in sheet_names:
                    missing_sheets.append(self.previous_sheet)

                if missing_sheets:
                    error_msg = f"The Excel file does not contain the required sheet(s): {', '.join(missing_sheets)}.\n"
                    error_msg += f"The Excel file contains the following sheets: {', '.join(sheet_names)}."
                    raise ValueError(f"Required sheets not found: {error_msg}")

                return original_init(self, excel_file, first_year, gui_year, non_current_assets, current_assets,
                            current_liabilities, non_current_liabilities, equity, revenue_items,
                            cost_of_sales_items, closing_inventories, other_income_items,
                            general_admin_expenses_items, finance_costs_items, tax_items)
            except ValueError as e:
                raise
            except Exception as e:
                raise

        DataLoader.__init__ = custom_init

        try:
            result, error_message = self._document_generator.generate_document(
                business_type=self.business_type.get(),
                excel_file=excel_file_path,
                output_path=output_path,
                current_year=current_year,
                first_year=self.first_year.get(),
                company_name_en=self.company_name_en.get(),
                company_name_cn=self.company_name_cn.get(),
                company_address=self.company_address.get(),
                business_description=self.business_description.get(),
                additional_business_description=self.additional_business_description.get(),
                last_day_of_year=self.last_day_of_year.get(),
                date_of_incorporation=self.date_of_incorporation.get(),
                audit_firm=self.audit_firm.get(),
                approval_date=self.approval_date.get(),
                auditor_name=self.auditor_name.get(),
                auditor_license=self.auditor_license.get(),
                currency=self.currency.get(),
                currency_desc=self.currency_desc.get(),
                currency_full_desc=self.currency_full_desc.get(),
                directors=self.directors.get().splitlines(),
                shareholders=self.shareholders.get().splitlines(),
                shares_curr=self.shares_curr.get(),
                shares_prev=self.shares_prev.get(),
                non_current_assets=self.non_current_assets,
                current_assets=self.current_assets,
                current_liabilities=self.current_liabilities,
                non_current_liabilities=self.non_current_liabilities,
                equity=self.equity,
                revenue_items=self.revenue_items,
                cost_of_sales_items=self.cost_of_sales_items,
                closing_inventories=self.closing_inventories,
                other_income_items=self.other_income_items,
                general_admin_expenses_items=self.general_admin_expenses_items,
                finance_costs_items=self.finance_costs_items,
                tax_items=self.tax_items,
                has_name_changed=self.has_name_changed.get(),
                passed_date=self.passed_date.get(),
                new_company_name=self.new_company_name.get(),
                effective_date=self.effective_date.get(),
                old_company_name=self.old_company_name.get(),
                has_related_party=self.has_related_party.get(),
                inventory_valuation=self.inventory_valuation.get(),
                tax_opt=self.tax_opt.get(),
                capital_increase=self.capital_increase.get(),
                has_ultimate_company=self.has_ultimate_company.get(),
                ultimate_company_option=self.ultimate_company_option.get(),
                ultimate_company_name1=self.ultimate_company_name1.get(),
                ultimate_company_location1=self.ultimate_company_location1.get(),
                ultimate_company_name2=self.ultimate_company_name2.get(),
                ultimate_company_location2=self.ultimate_company_location2.get(),
                investment_in_company=self.investment_in_company.get(),
                investment_in_security=self.investment_in_security.get(),
                audit_opinion=self.audit_opinion.get(),
                audit_type=self.audit_type.get(),
            )
            if result is None:
                self.status_label.config(text=error_message)
                messagebox.showwarning("Warning", error_message)
                return

            if error_message:
                messagebox.showwarning("Warning", error_message)

            if not os.path.exists(output_path):
                default_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "audit_report_filled.docx")
                if os.path.exists(default_path):
                    import shutil
                    shutil.copy2(default_path, output_path)
                    print(f"Moved file from {default_path} to {output_path}")

            if not os.path.exists(output_path):
                self.show_error(f"Generated file not found at {output_path}", "Error: Generated file not found")
                return

            self.status_label.config(text=f"Report generated successfully: {output_path}")

            if messagebox.askyesno("Success", f"Report generated successfully at {output_path}. Would you like to open it now?"):
                if sys.platform == 'darwin':
                    os.system(f"open '{output_path}'")
                elif sys.platform == 'win32':
                    os.system(f'start "" "{output_path}"')
                else:
                    os.system(f"xdg-open '{output_path}'")

        except UnrecognizedItemError as e:
            self.status_label.config(text="Error: Unrecognized item in TB sheet")
            messagebox.showwarning("Warning", str(e))
            return

        except InvalidTBSheetFormatError as e:
            self.status_label.config(text="Error: Invalid TB sheet format")
            messagebox.showwarning("Warning", str(e))
            return

        except ValueError as e:
            if "Required sheets not found" in str(e) or "Required sheet not found" in str(e):
                self.status_label.config(text="Error: Missing required sheets")
            else:
                self.show_error(f"Failed to generate report: {str(e)}")
                import traceback
                traceback.print_exc()
            return

        except Exception as e:
            self.show_error(f"Failed to generate report: {str(e)}")
            import traceback
            traceback.print_exc()
            return

        finally:
            DataLoader.__init__ = original_init

if __name__ == "__main__":
    app = AuditReportGUI()
    app.mainloop()