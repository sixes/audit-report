import tkinter as tk
from tkinter import ttk, messagebox
from .gui_utils import center_window

class ManageCategoriesDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Manage Categories")
        self.geometry("800x600")
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        # Center the dialog
        center_window(self, parent)

        # Use parent's category manager
        self.category_manager = parent.category_manager
        self.categories = {
            'Non-Current Assets': self.category_manager.categories['non_current_assets'],
            'Current Assets': self.category_manager.categories['current_assets'],
            'Current Liabilities': self.category_manager.categories['current_liabilities'],
            'Non-Current Liabilities': self.category_manager.categories['non_current_liabilities'],
            'Equity': self.category_manager.categories['equity'],
            'Revenue Items': self.category_manager.categories['revenue_items'],
            'Cost of Sales Items': self.category_manager.categories['cost_of_sales_items'],
            'Closing Inventories': self.category_manager.categories['closing_inventories'],
            'Other Income Items': self.category_manager.categories['other_income_items'],
            'General Admin Expenses Items': self.category_manager.categories['general_admin_expenses_items'],
            'Finance Costs Items': self.category_manager.categories['finance_costs_items'],
            'Tax Items': self.category_manager.categories['tax_items']
        }

        # Main container
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Treeview for categories and items
        self.tree_frame = ttk.LabelFrame(self.main_frame, text="Categories and Items")
        self.tree_frame.pack(fill='both', expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(self.tree_frame, columns=('Item',), show='tree headings')
        self.tree.heading('Item', text='Item')
        self.tree.column('Item', width=400)
        self.tree.pack(fill='both', expand=True, padx=5, pady=5)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)

        # Populate tree
        self.populate_tree()

        # Input frame for adding/modifying items
        self.input_frame = ttk.LabelFrame(self.main_frame, text="Manage Items")
        self.input_frame.pack(fill='x', padx=5, pady=5)

        # Category selection
        ttk.Label(self.input_frame, text="Category:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.category_var = tk.StringVar()
        self.category_combobox = ttk.Combobox(self.input_frame, textvariable=self.category_var, state='readonly')
        self.category_combobox['values'] = list(self.categories.keys())
        self.category_combobox.set('Current Assets')
        self.category_combobox.grid(row=0, column=1, sticky='ew', padx=5, pady=5)

        # Item entry
        ttk.Label(self.input_frame, text="Item Name:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.item_entry = ttk.Entry(self.input_frame)
        self.item_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)

        # Action buttons
        self.action_frame = ttk.Frame(self.input_frame)
        self.action_frame.grid(row=2, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        self.add_btn = ttk.Button(self.action_frame, text="Add Item", command=self.add_item)
        self.add_btn.pack(side='left', padx=5)

        self.modify_btn = ttk.Button(self.action_frame, text="Modify Item", command=self.modify_item, state='disabled')
        self.modify_btn.pack(side='left', padx=5)

        self.delete_btn = ttk.Button(self.action_frame, text="Delete Item", command=self.delete_item, state='disabled')
        self.delete_btn.pack(side='left', padx=5)

        # Status bar
        self.status_var = tk.StringVar(value="Select a category and enter an item to manage.")
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status_var, relief='sunken')
        self.status_label.pack(fill='x', padx=5, pady=5)

        # Bottom buttons
        self.buttons_frame = ttk.Frame(self.main_frame)
        self.buttons_frame.pack(fill='x', padx=5, pady=5)

        ttk.Button(self.buttons_frame, text="Save and Close", command=self.save_and_close).pack(side='right', padx=5)
        ttk.Button(self.buttons_frame, text="Cancel", command=self.destroy).pack(side='right', padx=5)

        # Configure grid weights
        self.input_frame.columnconfigure(1, weight=1)

    def populate_tree(self):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Add categories and their items
        for category, items in self.categories.items():
            category_id = self.tree.insert('', 'end', text=category, open=True)
            for item in items:  # All categories are lists
                self.tree.insert(category_id, 'end', text=item)

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if not selected:
            self.modify_btn.config(state='disabled')
            self.delete_btn.config(state='disabled')
            self.item_entry.delete(0, tk.END)
            self.status_var.set("Select an item to modify or delete.")
            return

        selected_item = selected[0]
        parent = self.tree.parent(selected_item)
        if not parent:  # Category selected
            self.modify_btn.config(state='disabled')
            self.delete_btn.config(state='disabled')
            self.item_entry.delete(0, tk.END)
            self.status_var.set("Select an item to modify or delete.")
            return

        # Item selected
        item_text = self.tree.item(selected_item, 'text')
        parent_text = self.tree.item(parent, 'text')
        self.category_var.set(parent_text)
        self.item_entry.delete(0, tk.END)
        self.item_entry.insert(0, item_text)
        self.modify_btn.config(state='normal')
        self.delete_btn.config(state='normal')
        self.status_var.set(f"Selected item: {item_text}")

    def show_error(self, message, status_text=None):
        """Show an error message and update status."""
        messagebox.showerror("Error", message)
        self.status_var.set(status_text or f"Error: {message}")

    def add_item(self):
        new_item = self.item_entry.get().strip()
        if not new_item:
            self.show_error("Please enter an item name.")
            return

        selected_category = self.category_var.get()
        category_key = {
            'Non-Current Assets': 'non_current_assets',
            'Current Assets': 'current_assets',
            'Current Liabilities': 'current_liabilities',
            'Non-Current Liabilities': 'non_current_liabilities',
            'Equity': 'equity',
            'Revenue Items': 'revenue_items',
            'Cost of Sales Items': 'cost_of_sales_items',
            'Closing Inventories': 'closing_inventories',
            'Other Income Items': 'other_income_items',
            'General Admin Expenses Items': 'general_admin_expenses_items',
            'Finance Costs Items': 'finance_costs_items',
            'Tax Items': 'tax_items'
        }[selected_category]

        try:
            self.category_manager.add_item(category_key, new_item)
            self.populate_tree()
            self.item_entry.delete(0, tk.END)
            self.status_var.set(f"Item '{new_item}' added to {selected_category}.")
        except ValueError as e:
            self.show_error(str(e))

    def modify_item(self):
        selected = self.tree.selection()
        if not selected:
            self.show_error("Please select an item to modify.")
            return

        selected_item = selected[0]
        if not self.tree.parent(selected_item):  # Category selected
            self.show_error("Please select an item, not a category.")
            return

        new_item = self.item_entry.get().strip()
        if not new_item:
            self.show_error("Please enter a new item name.")
            return

        parent = self.tree.parent(selected_item)
        selected_category = self.tree.item(parent, 'text')
        old_item = self.tree.item(selected_item, 'text')
        category_key = {
            'Non-Current Assets': 'non_current_assets',
            'Current Assets': 'current_assets',
            'Current Liabilities': 'current_liabilities',
            'Non-Current Liabilities': 'non_current_liabilities',
            'Equity': 'equity',
            'Revenue Items': 'revenue_items',
            'Cost of Sales Items': 'cost_of_sales_items',
            'Closing Inventories': 'closing_inventories',
            'Other Income Items': 'other_income_items',
            'General Admin Expenses Items': 'general_admin_expenses_items',
            'Finance Costs Items': 'finance_costs_items',
            'Tax Items': 'tax_items'
        }[selected_category]

        try:
            self.category_manager.modify_item(category_key, old_item, new_item)
            self.populate_tree()
            self.item_entry.delete(0, tk.END)
            self.modify_btn.config(state='disabled')
            self.delete_btn.config(state='disabled')
            self.status_var.set(f"Item modified to '{new_item}' in {selected_category}.")
        except ValueError as e:
            self.show_error(str(e))

    def delete_item(self):
        selected = self.tree.selection()
        if not selected:
            self.show_error("Please select an item to delete.")
            return

        selected_item = selected[0]
        if not self.tree.parent(selected_item):  # Category selected
            self.show_error("Please select an item, not a category.")
            return

        parent = self.tree.parent(selected_item)
        selected_category = self.tree.item(parent, 'text')
        item_text = self.tree.item(selected_item, 'text')
        category_key = {
            'Non-Current Assets': 'non_current_assets',
            'Current Assets': 'current_assets',
            'Current Liabilities': 'current_liabilities',
            'Non-Current Liabilities': 'non_current_liabilities',
            'Equity': 'equity',
            'Revenue Items': 'revenue_items',
            'Cost of Sales Items': 'cost_of_sales_items',
            'Closing Inventories': 'closing_inventories',
            'Other Income Items': 'other_income_items',
            'General Admin Expenses Items': 'general_admin_expenses_items',
            'Finance Costs Items': 'finance_costs_items',
            'Tax Items': 'tax_items'
        }[selected_category]

        self.category_manager.delete_item(category_key, item_text)
        self.populate_tree()
        self.item_entry.delete(0, tk.END)
        self.modify_btn.config(state='disabled')
        self.delete_btn.config(state='disabled')
        self.status_var.set(f"Item '{item_text}' deleted from {selected_category}.")

    def save_and_close(self):
        self.category_manager.save()
        # Update parent attributes
        self.parent.non_current_assets = self.category_manager.categories['non_current_assets']
        self.parent.current_assets = self.category_manager.categories['current_assets']
        self.parent.current_liabilities = self.category_manager.categories['current_liabilities']
        self.parent.non_current_liabilities = self.category_manager.categories['non_current_liabilities']
        self.parent.equity = self.category_manager.categories['equity']
        self.parent.revenue_items = self.category_manager.categories['revenue_items']
        self.parent.cost_of_sales_items = self.category_manager.categories['cost_of_sales_items']
        self.parent.closing_inventories = self.category_manager.categories['closing_inventories']
        self.parent.other_income_items = self.category_manager.categories['other_income_items']
        self.parent.general_admin_expenses_items = self.category_manager.categories['general_admin_expenses_items']
        self.parent.finance_costs_items = self.category_manager.categories['finance_costs_items']
        self.parent.tax_items = self.category_manager.categories['tax_items']
        self.destroy()