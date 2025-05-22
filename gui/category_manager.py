import json
from .gui_utils import load_categories

class CategoryManager:
    def __init__(self, config_file='categories.json'):
        self.config_file = config_file
        self.categories = self.load_default_categories()
        self.load_from_file()

    def load_default_categories(self):
        """Return default category lists."""
        return {
            'non_current_assets': [item.lower() for item in [
                'Intangible assets', 'Investments in a subsidiary', 'Investments in an associate',
                'Investments in subsidiaries', 'Interests in subsidiaries', 'Interests in a subsidiary', 'Investments in associates',
                'Long-term investments', 'Property, plant and equipment', 'Deferred tax assets'
            ]],
            'current_assets': [item.lower() for item in [
                'Amount due from director', 'Amount due from a director', 'Amount due from the director', 'Amount due from directors',
                'Amount due from shareholder', 'Amount due from a shareholder', 'Amount due from the shareholder', 'Amount due from shareholders',
                'Amount due from final holding parent company', 'Amount due from a final holding parent company', 'Amount due from the final holding parent company', 'Amount due from final holding parent companies',
                'Amount due from immediate parent company', 'Amount due from an immediate parent company', 'Amount due from the immediate parent company', 'Amount due from immediate parent companies',
                'Amount due from holding company', 'Amount due from a holding company', 'Amount due from the holding company', 'Amount due from holding companies', 
                'Amount due from ultimate holding company', 'Amount due from an ultimate holding company', 'Amount due from the ultimate holding company', 'Amount due from ultimate holding companies',
                'Amount due from related company', 'Amount due from a related company', 'Amount due from the related company', 'Amount due from related companies',
                'Cash and bank balances', 'Cash and cash equivalents', 'Current investments',
                'Inventories', 'Other receivables', 'Prepayments', 'Prepayment','Rental deposit', 'Tax recoverable',
                'Accounts receivable', 'Trade receivables', 'Prepaid expenses'
            ]],
            'current_liabilities': [item.lower() for item in [
                'Accrued expenses', 
                'Amount due to director', 'Amount due to a director', 'Amount due to the director', 'Amount due to directors',
                'Amount due to subsidiary company', 'Amount due to a subsidiary company', 'Amount due to the subsidiary company', 'Amount due to subsidiary companies',
                'Amount due to related company', 'Amount due to a related company', 'Amount due to the related company', 'Amount due to related companies',
                'Amount due to shareholder', 'Amount due to a shareholder', 'Amount due to the shareholder', 'Amount due to shareholders',    
                'Amount due to final holding parent company', 'Amount due to a final holding parent company', 'Amount due to the final holding parent company', 'Amount due to final holding parent companies',
                'Amount due to immediate parent company', 'Amount due to an immediate parent company', 'Amount due to the immediate parent company', 'Amount due to immediate parent companies',    
                'Amount due to holding company', 'Amount due to a holding company', 'Amount due to the holding company', 'Amount due to holding companies',
                'Amount due to ultimate holding company', 'Amount due to an ultimate holding company', 'Amount due to the ultimate holding company', 'Amount due to ultimate holding companies',
                'Bank overdraft', 'Borrowings-secured', 'Deposits received', 'Deposit received',
                'Other payables', 'Short-term borrowings', 'Tax payable',
                'Trade payables', 'Accounts payable', 'Accounts payables','Account payables','Accrued wages',
                'Accounts and other payables'
            ]],
            'non_current_liabilities': [item.lower() for item in [
                'Deferred tax liabilities', 'Obligations under finance leases', 'Long-term borrowings'
            ]],
            'equity': [item.lower() for item in [
                'Capital reserves', 'Share capital', 'Reserves', 'Dividends paid to shareholders', 'Dividends paid to a shareholder'
            ]],
            'revenue_items': [item.lower() for item in [
                'Sales of goods', 'Services fee income', 'Agency fee income'
            ]],
            'cost_of_sales_items': [item.lower() for item in [
                'Direct costs', 'Cost of services', 'Opening inventories', 'Purchases'
            ]],
            'closing_inventories': [item.lower() for item in ['Closing inventories']],  # Normalized to list
            'other_income_items': [item.lower() for item in [
                'Bank interest income', 'Commission income', 'Dividend income', 'Exchange gains', 'Exchange gain',
                'Gains on disposal of financial assets', 'Gains on fair value of investment securities',
                'Government grants', 'Reversal of impairment of investment',
                'Reversal of impairment losses on long-term investments', 'Sundry income', 'Refund of postage', 'Other income'
            ]],
            'general_admin_expenses_items': [item.lower() for item in [
                'Accountancy fee', 'Accounting fee', 'Advertising fee', 'Amortisation', 'Annual return fee', 'Audit fee',
                "Auditors' remuneration", 'Bank charges', 'Bank charges.', 'Bank charges and interest', 'Building management fee', 'Business registration fee',
                'Business trips expenses', 'Commission', 'Compensation', 'Conference expenses',
                'Consulting fee', 'Declaration', 'Depreciation', 'Director\'s remuneration', 'Director’s remuneration',
                'Entertainment', 'Exchange loss', 'Exchange losses', 'Exhibition fee', 'Exhibition fees', 'FBA operation fee', 'FBA storage fee',
                'Filing fees',
                'Impairment loss on long-term investment', 'Impairment losses on long-term investments',
                'Impairment of investments in a subsidiary', 'Impairment of investments in subsidiaries',
                'Impairment loss for trade and other receivables', 'Impairment loss for trade receivables',
                'Impairment loss for other receivables', 
                'Inspection fee', 'Insurance', 'Insurance expenses',
                'Inventories write down', 'Legal and professional fee', 'Legal and professional fees',
                'Loss on disposal of financial assets', 'Losses on disposal of financial assets',
                'Loss on fair value of investment securities', 'Losses on fair value of investments in securities',
                'Management fee', 'Material fee', 'Motor car expenses',
                'MPF', 'MPF contribution', 'Network service charges', 'Office supplies', 'Office expenses',
                'Operation fee', 'Packing fee', 'Penalty', 'Platform commission fee', 'Platform fee',
                'Platform outsourcing management fee', 'Postage and courier', 'Postage',
                'Preliminary expenses', 'Printing and stationery', 'Promotion fee',
                'Provision for bad debts', 'Rent and rates', 'Repair expenses', 'Rental vehicles',
                'Salaries', 'Sample charges', 'Secretarial fee', 'Staff welfare', 'Welfare fee',
                'Storage fee', 'Stamp duty', 'Technical service fee', 'Technical services fee',
                'Telephone charges', 'Training fee', 'Transportation fee', 'Travelling',
                'Value-added tax', 'Water and electricity', 'Asset impairment loss', 'Communication fee', 'Customs fee', 'Services fee',
                'Sundry expenses', 'Travel expenses', 'Maintenance fee'
            ]],
            'finance_costs_items': [item.lower() for item in [
                'Loan interest'
            ]],
            'tax_items': [item.lower() for item in [
                'Tax payable', 'Tax recoverable'
            ]]
        }

    def load_from_file(self):
        """Load categories from JSON file, updating defaults."""
        defaults = self.categories
        loaded = load_categories(self.config_file, defaults)
        for key in defaults:
            self.categories[key] = loaded.get(key, defaults[key])
        # Ensure closing_inventories is a list
        if isinstance(self.categories['closing_inventories'], str):
            self.categories['closing_inventories'] = [self.categories['closing_inventories']]

    def add_item(self, category, item):
        """Add an item to a category."""
        if item in self.categories[category]:
            raise ValueError("Item already exists in this category")
        self.categories[category].append(item.lower())

    def modify_item(self, category, old_item, new_item):
        """Modify an existing item in a category."""
        items = self.categories[category]
        if new_item in items and new_item != old_item:
            raise ValueError("New item name already exists in this category")
        index = items.index(old_item)
        items[index] = new_item.lower()

    def delete_item(self, category, item):
        """Delete an item from a category."""
        self.categories[category].remove(item)
        if category == 'closing_inventories' and not self.categories[category]:
            self.categories[category] = ['Closing inventories']

    def save(self):
        """Save categories to JSON file."""
        with open(self.config_file, 'w') as f:
            json.dump(self.categories, f, indent=4)

    def get_categories(self):
        """Return the categories dictionary."""
        return self.categories