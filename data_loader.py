import pandas as pd
import re
from exceptions import InvalidTBSheetFormatError, InvalidItemNameError, UnrecognizedItemError

class DataLoader:
    def __init__(self, excel_file, first_year, current_year, non_current_assets, current_assets, current_liabilities,
                 non_current_liabilities, equity, revenue_items, cost_of_sales_items, closing_inventories,
                 other_income_items, general_admin_expenses_items, finance_costs_items, tax_items):
        self.excel_file = excel_file
        self.first_year = first_year
        self.current_year = current_year
        self.previous_year = current_year - 1
        self.current_sheet = f"{self.current_year}TB"
        self.previous_sheet = f"{self.previous_year}TB"
        # Store category lists
        self.non_current_assets = non_current_assets
        self.current_assets = current_assets
        self.current_liabilities = current_liabilities
        self.non_current_liabilities = non_current_liabilities
        self.equity = equity
        self.revenue_items = revenue_items
        self.cost_of_sales_items = cost_of_sales_items
        self.closing_inventories = closing_inventories
        self.other_income_items = other_income_items
        self.general_admin_expenses_items = general_admin_expenses_items
        self.finance_costs_items = finance_costs_items
        self.tax_items = tax_items
        self.data = self._load_data()

    def _load_data(self):
        try:
            xl = pd.ExcelFile(self.excel_file)
            sheet_names = xl.sheet_names
            data = {}

            # Regular expression to match only letters and spaces
            valid_item_name_pattern = r'^[a-zA-Z\s\/\-,\.\']+$'

            # Helper function to validate item names
            def validate_item_name(item, sheet_name):
                if pd.isna(item) or not str(item).strip() or str(item).strip() in ('合計', '董事簽名：'):
                    return
                item_str = str(item).strip()
                if not re.match(valid_item_name_pattern, item_str):
                    raise InvalidItemNameError(
                        f"Invalid item name '{item_str}' in sheet '{sheet_name}'. "
                        "Item names must contain only letters and spaces."
                    )

            # Load current year TB
            if self.current_sheet not in sheet_names:
                raise ValueError(f"Sheet {self.current_sheet} not found in Excel file")
            current_df = pd.read_excel(xl, sheet_name=self.current_sheet, header=None, skiprows=3)
            if len(current_df.columns) < 3:
                raise InvalidTBSheetFormatError(
                    "Failed to recognize the sheets. The first 3 rows are the headers, "
                    "the data should start from row 4 with columns 'Item', 'Debtor', 'Creditor'."
                )
            current_df.columns = ['Item', 'Debtor', 'Creditor']
            # Normalize Item column: strip whitespace and convert to lowercase
            current_df['Item'] = current_df['Item'].astype(str).str.strip().str.lower()
            current_df = current_df[current_df['Item'].notna() & (current_df['Item'] != "") & (current_df['Item'] != "nan")]
            current_df[['Debtor', 'Creditor']] = current_df[['Debtor', 'Creditor']].fillna(0)
            # Validate item names for current year
            for item in current_df['Item']:
                validate_item_name(item, self.current_sheet)
            try:
                current_df['Debtor'] = pd.to_numeric(current_df['Debtor'], errors='raise').round(2)
                current_df['Creditor'] = pd.to_numeric(current_df['Creditor'], errors='raise').round(2)
            except ValueError as e:
                raise InvalidTBSheetFormatError(
                    "Failed to recognize the sheets. The first 3 rows are the headers, "
                    "the data should start from row 4 with columns 'Item', 'Debtor', 'Creditor'. "
                    f"Error in data conversion: {str(e)}"
                )
            current_df = current_df.fillna(0)
            data[self.current_year] = current_df

            if self.first_year:
                previous_df = pd.DataFrame(columns=['Item', 'Debtor', 'Creditor'])
                data[self.previous_year] = previous_df
                return data
            # Load previous year TB
            if self.previous_sheet not in sheet_names:
                raise ValueError(f"Sheet {self.previous_sheet} not found in Excel file")
            previous_df = pd.read_excel(xl, sheet_name=self.previous_sheet, header=None, skiprows=3)
            if len(previous_df.columns) < 3:
                raise InvalidTBSheetFormatError(
                    "Failed to recognize the sheets. The first 3 rows are the headers, "
                    "the data should start from row 4 with columns 'Item', 'Debtor', 'Creditor'."
                )
            previous_df.columns = ['Item', 'Debtor', 'Creditor']
            # Normalize Item column: strip whitespace and convert to lowercase
            previous_df['Item'] = previous_df['Item'].astype(str).str.strip().str.lower()
            previous_df = previous_df[previous_df['Item'].notna() & (previous_df['Item'] != "") & (previous_df['Item'] != "nan")]
            previous_df[['Debtor', 'Creditor']] = previous_df[['Debtor', 'Creditor']].fillna(0)
            # Validate item names for previous year
            for item in previous_df['Item']:
                validate_item_name(item, self.previous_sheet)
            try:
                previous_df['Debtor'] = pd.to_numeric(previous_df['Debtor'], errors='raise').round(2)
                previous_df['Creditor'] = pd.to_numeric(previous_df['Creditor'], errors='raise').round(2)
            except ValueError as e:
                raise InvalidTBSheetFormatError(
                    "Failed to recognize the sheets. The first 3 rows are the headers, "
                    "the data should start from row 4 with columns 'Item', 'Debtor', 'Creditor'. "
                    f"Error in data conversion: {str(e)}"
                )
            previous_df = previous_df.fillna(0)
            data[self.previous_year] = previous_df

            return data
        except InvalidTBSheetFormatError as e:
            raise e
        except InvalidItemNameError as e:
            raise e
        except Exception as e:
            raise Exception(f"Failed to load Excel file: {str(e)}")

    def _get_balance_before_period(self, year):
        """Extract the 'Balance before current period' item from the TB sheet for a given year.
        Handles both positive (Creditor) and negative (Debtor) values."""
        if year not in self.data:
            return 0
        df = self.data[year]
        balance_before_items = ['balance before current period', 'balance bf current period', 'balance b/f current period']
        for _, row in df.iterrows():
            item = row['Item']
            if item in balance_before_items:
                creditor = float(row['Creditor'] or 0)
                debtor = float(row['Debtor'] or 0)
                return creditor if creditor != 0 else -debtor
        return 0

    def _categorize_items(self, year):
        df = self.data[year]

        # Use category lists passed to the constructor
        non_current_assets = self.non_current_assets
        current_assets = self.current_assets
        current_liabilities = self.current_liabilities
        non_current_liabilities = self.non_current_liabilities
        equity = self.equity
        revenue_items = self.revenue_items
        cost_of_sales_items = self.cost_of_sales_items
        closing_inventories = self.closing_inventories
        other_income_items = self.other_income_items
        general_admin_expenses_items = self.general_admin_expenses_items
        finance_costs_items = self.finance_costs_items
        tax_items = self.tax_items

        # Normalize lists for case-insensitive comparison
        non_current_assets_lower = [item.lower() for item in non_current_assets]
        current_assets_lower = [item.lower() for item in current_assets]
        current_liabilities_lower = [item.lower() for item in current_liabilities]
        non_current_liabilities_lower = [item.lower() for item in non_current_liabilities]
        equity_lower = [item.lower() for item in equity]
        revenue_items_lower = [item.lower() for item in revenue_items]
        cost_of_sales_items_lower = [item.lower() for item in cost_of_sales_items]
        closing_inventories_lower = closing_inventories.lower() if isinstance(closing_inventories, str) else closing_inventories[0].lower()
        other_income_items_lower = [item.lower() for item in other_income_items]
        general_admin_expenses_items_lower = [item.lower() for item in general_admin_expenses_items]
        finance_costs_items_lower = [item.lower() for item in finance_costs_items]
        tax_items_lower = [item.lower() for item in tax_items]
        balance_before_items_lower = ['balance before current period', 'balance bf current period', 'balance b/f current period']

        all_valid_items = (
            non_current_assets + current_assets + current_liabilities + non_current_liabilities + equity +
            revenue_items + cost_of_sales_items + ([closing_inventories] if isinstance(closing_inventories, str) else closing_inventories) +
            other_income_items + general_admin_expenses_items + finance_costs_items + tax_items +
            ['balance before current period', 'balance bf current period', 'balance b/f current period']
        )
        all_valid_items_lower = [item.lower() for item in all_valid_items]

        revenue = 0
        cost_of_sales = 0
        closing_inv = 0
        other_income = 0
        general_admin_expenses = 0
        finance_costs = 0
        taxation = 0

        revenue_items_details = []
        cost_items_details = []
        other_income_details = []
        general_admin_expenses_details = []

        balance_sheet = {
            'non_current_assets': [],
            'current_assets': [],
            'current_liabilities': [],
            'non_current_liabilities': [],
            'equity': [],
            'total_non_current_assets': 0,
            'total_current_assets': 0,
            'total_current_liabilities': 0,
            'total_non_current_liabilities': 0,
            'total_equity': 0,
            'net_assets': 0
        }

        for _, row in df.iterrows():
            item = row['Item']
            debtor = float(row['Debtor'] or 0)
            creditor = float(row['Creditor'] or 0)

            if item in ['董事簽名：', '合計', '0', 'taxation']:
                continue

            if item not in all_valid_items_lower:
                print(f'item: {item}')
                print(f'all_valid_items_lower: {all_valid_items_lower}')
                raise UnrecognizedItemError(f"Unrecognized item found in TB sheet: '{item}'")

            if item in balance_before_items_lower:
                continue

            if item in revenue_items_lower:
                revenue += creditor
                if creditor != 0:
                    revenue_items_details.append({
                        'name': item,
                        'value': creditor
                    })
            elif item in cost_of_sales_items_lower:
                cost_of_sales += debtor
                if debtor != 0:
                    cost_items_details.append({
                        'name': item,
                        'value': debtor
                    })
            elif item == closing_inventories_lower:
                closing_inv += debtor
                if debtor != 0:
                    cost_items_details.append({
                        'name': item,
                        'value': -debtor
                    })
            elif item in other_income_items_lower:
                other_income += creditor
                if creditor != 0:
                    other_income_details.append({
                        'name': item,
                        'value': creditor
                    })
            elif item in general_admin_expenses_items_lower:
                general_admin_expenses += debtor
                if debtor != 0:
                    general_admin_expenses_details.append({
                        'name': item,
                        'value': debtor
                    })
            elif item in finance_costs_items_lower:
                finance_costs += debtor

            if item in non_current_assets_lower:
                idx = non_current_assets_lower.index(item)
                original_item = non_current_assets[idx]
                value = debtor
                balance_sheet['non_current_assets'].append({
                    'name': item,
                    'value': value
                })
                balance_sheet['total_non_current_assets'] += value
            elif item in current_assets_lower:
                idx = current_assets_lower.index(item)
                original_item = current_assets[idx]
                value = debtor
                balance_sheet['current_assets'].append({
                    'name': item,
                    'value': value
                })
                balance_sheet['total_current_assets'] += value
            elif item in current_liabilities_lower:
                idx = current_liabilities_lower.index(item)
                original_item = current_liabilities[idx]
                value = creditor
                balance_sheet['current_liabilities'].append({
                    'name': item,
                    'value': value
                })
                balance_sheet['total_current_liabilities'] += value
            elif item in non_current_liabilities_lower:
                idx = non_current_liabilities_lower.index(item)
                original_item = non_current_liabilities[idx]
                value = creditor
                balance_sheet['non_current_liabilities'].append({
                    'name': item,
                    'value': value
                })
                balance_sheet['total_non_current_liabilities'] += value
            elif item in equity_lower:
                idx = equity_lower.index(item)
                original_item = equity[idx]
                value = creditor - debtor
                balance_sheet['equity'].append({
                    'name': item,
                    'value': value
                })
                balance_sheet['total_equity'] += value

        # Second pass for 'taxation' item
        for _, row in df.iterrows():
            item = row['Item']
            debtor = float(row['Debtor'] or 0)
            creditor = float(row['Creditor'] or 0)
            if item == 'taxation':
                taxation = creditor - debtor
                break

        cost_of_sales += closing_inv
        gross_profit = revenue - cost_of_sales
        calc_total = gross_profit + other_income
        profit_before_tax = calc_total - general_admin_expenses - finance_costs
        profit_for_year = profit_before_tax + taxation

        balance_sheet['net_assets'] = (
            balance_sheet['total_non_current_assets'] +
            balance_sheet['total_current_assets'] -
            balance_sheet['total_current_liabilities'] -
            balance_sheet['total_non_current_liabilities']
        )

        return {
            "Revenue": revenue,
            "CostOfSales": cost_of_sales,
            "GrossProfit": gross_profit,
            "OtherIncome": other_income,
            "GeneralAdminExpenses": general_admin_expenses,
            "FinanceCosts": finance_costs,
            "CalcTotal": calc_total,
            "ProfitBeforeTax": profit_before_tax,
            "Taxation": taxation,
            "ProfitForYear": profit_for_year,
            "BalanceSheet": balance_sheet,
            "RevenueItemsDetails": revenue_items_details,
            "CostItemsDetails": cost_items_details,
            "OtherIncomeDetails": other_income_details,
            "GeneralAdminExpensesDetails": general_admin_expenses_details
        }

    def get_income_statement(self, year):
        if year not in [self.current_year, self.previous_year]:
            raise ValueError(f"Year must be {self.current_year} or {self.previous_year}")
        return self._categorize_items(year)