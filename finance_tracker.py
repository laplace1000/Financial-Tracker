import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from datetime import datetime
import os
import json

class FinanceTracker:
    def __init__(self, file_path='finance_tracker.xlsx'):
        self.file_path = file_path
        self.sheets = ['Income', 'Expenses', 'Savings']
        self.budget_file = 'budgets.json'  # Store budgets in a separate JSON file
        
        # Initialize budget storage
        self._init_budget_file()
        self._initialize_excel_file()
    
    def _initialize_excel_file(self):
        """Initialize Excel file with required sheets if it doesn't exist or is missing sheets."""
        try:
            if not Path(self.file_path).exists():
                # Create new file with all sheets
                dfs = {sheet: pd.DataFrame(columns=['Month', 'Amount', 'Category', 'Notes']) 
                      for sheet in self.sheets}
                
                with pd.ExcelWriter(self.file_path, engine='openpyxl') as writer:
                    for sheet_name, df in dfs.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # Check existing sheets
                try:
                    existing_sheets = pd.ExcelFile(self.file_path, engine='openpyxl').sheet_names
                except:
                    existing_sheets = []
                
                missing_sheets = set(self.sheets) - set(existing_sheets)
                
                if missing_sheets:
                    # Read existing sheets
                    all_data = {}
                    for sheet in existing_sheets:
                        try:
                            all_data[sheet] = pd.read_excel(self.file_path, 
                                                          sheet_name=sheet, 
                                                          engine='openpyxl')
                        except:
                            continue
                    
                    # Create missing sheets
                    for sheet in missing_sheets:
                        all_data[sheet] = pd.DataFrame(columns=['Month', 'Amount', 'Category', 'Notes'])
                    
                    # Write all sheets back
                    with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='w') as writer:
                        for sheet_name, df in all_data.items():
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # Write any remaining missing sheets
                        for sheet in missing_sheets - set(all_data.keys()):
                            empty_df = pd.DataFrame(columns=['Month', 'Amount', 'Category', 'Notes'])
                            empty_df.to_excel(writer, sheet_name=sheet, index=False)
            
            # Verify all sheets exist
            existing_sheets = pd.ExcelFile(self.file_path, engine='openpyxl').sheet_names
            if not all(sheet in existing_sheets for sheet in self.sheets):
                raise ValueError("Failed to create all required sheets")
            
        except Exception as e:
            raise ValueError(f"Error initializing Excel file: {str(e)}")
    
    def add_entry(self, entry_type, amount, category, notes='', date=None):
        """Add a new financial entry."""
        if entry_type not in self.sheets:
            raise ValueError(f"Entry type must be one of {self.sheets}")
        
        # Use current month if date is not provided
        if date is None:
            date = datetime.now().strftime('%Y-%m')
        
        try:
            # Read all existing sheets first
            all_data = {}
            for sheet in self.sheets:
                try:
                    df = pd.read_excel(self.file_path, sheet_name=sheet)
                    # Ensure consistent data types
                    df['Month'] = pd.to_datetime(df['Month']).dt.strftime('%Y-%m')
                    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
                    df['Category'] = df['Category'].astype(str)
                    df['Notes'] = df['Notes'].astype(str)
                    all_data[sheet] = df
                except:
                    # If sheet doesn't exist, create empty DataFrame with specified dtypes
                    all_data[sheet] = pd.DataFrame({
                        'Month': pd.Series(dtype='str'),
                        'Amount': pd.Series(dtype='float64'),
                        'Category': pd.Series(dtype='str'),
                        'Notes': pd.Series(dtype='str')
                    })
            
            # Create new entry with explicit data types
            new_entry = pd.DataFrame({
                'Month': [date],
                'Amount': [float(amount)],
                'Category': [str(category)],
                'Notes': [str(notes)]
            })
            
            # Update the target sheet
            if all_data[entry_type].empty:
                all_data[entry_type] = new_entry
            else:
                all_data[entry_type] = pd.concat([all_data[entry_type], new_entry], ignore_index=True)
            
            # Write all sheets back to Excel in a single operation
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='w') as writer:
                for sheet_name, df in all_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            return True
        
        except Exception as e:
            raise ValueError(f"Error adding entry: {str(e)}")
    
    def get_monthly_summary(self, month=None):
        """Get financial summary for a specific month or all months."""
        summary = {}
        
        for sheet in self.sheets:
            df = pd.read_excel(self.file_path, sheet_name=sheet)
            if month:
                df = df[df['Month'] == month]
            summary[sheet] = df['Amount'].sum()
            
        return summary
    
    def visualize_monthly_data(self, months=6):
        """Create a bar chart showing financial data for the last n months."""
        data = []
        
        try:
            for sheet in self.sheets:
                df = pd.read_excel(self.file_path, sheet_name=sheet)
                if not df.empty:
                    # Convert Month to datetime and Amount to float
                    df['Month'] = pd.to_datetime(df['Month'])
                    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
                    
                    # Group by Month and sum Amount
                    monthly_totals = df.groupby('Month')['Amount'].sum().reset_index()
                    monthly_totals = monthly_totals.sort_values('Month', ascending=True).tail(months)
                    data.append(monthly_totals.assign(Type=sheet))
            
            if data:
                plot_data = pd.concat(data)
                
                plt.figure(figsize=(12, 6))
                sns.barplot(data=plot_data, x='Month', y='Amount', hue='Type')
                plt.title('Monthly Financial Summary')
                plt.xticks(rotation=45)
                plt.tight_layout()
                plt.show()
                
        except Exception as e:
            print(f"Error visualizing data: {str(e)}")

    def get_sheet_data(self, sheet_name):
        """Get processed data for a specific sheet."""
        try:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            if not df.empty:
                # Convert Month to datetime
                df['Month'] = pd.to_datetime(df['Month'])
                # Ensure Amount is float
                df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
                
                # Group by Month and sum Amount
                monthly_totals = df.groupby('Month')['Amount'].sum().reset_index()
                monthly_totals = monthly_totals.sort_values('Month', ascending=True).tail(6)
                return monthly_totals.assign(Type=sheet_name)
        except Exception as e:
            print(f"Error processing {sheet_name} data: {str(e)}")
        
        # Return empty DataFrame with correct structure if there's an error
        return pd.DataFrame(columns=['Month', 'Amount', 'Type'])

    def get_all_categories(self):
        """Get all unique categories across all sheets"""
        categories = set()
        try:
            for sheet in self.sheets:
                df = pd.read_excel(self.file_path, sheet_name=sheet)
                # Convert Category column to string and handle NaN values
                df['Category'] = df['Category'].fillna('').astype(str)
                # Only add non-empty categories
                categories.update([cat for cat in df['Category'].unique() if cat and cat != 'nan'])
            
            # Sort and return non-empty categories
            return sorted(list(categories)) if categories else []
        
        except Exception as e:
            print(f"Error getting categories: {str(e)}")
            return []

    def analyze_category(self, category):
        """Analyze spending/income for a specific category"""
        analysis = f"Analysis for category: {category}\n\n"
        
        for sheet in self.sheets:
            df = pd.read_excel(self.file_path, sheet_name=sheet)
            cat_data = df[df['Category'] == category]
            
            if not cat_data.empty:
                total = cat_data['Amount'].sum()
                avg = cat_data['Amount'].mean()
                analysis += f"{sheet}:\n"
                analysis += f"Total: ${total:,.2f}\n"
                analysis += f"Average: ${avg:,.2f}\n\n"
        
        return analysis

    def get_trending_categories(self):
        """Identify trending categories based on recent transactions"""
        trends = "Trending Categories (Last 3 months):\n\n"
        
        for sheet in self.sheets:
            df = pd.read_excel(self.file_path, sheet_name=sheet)
            df['Month'] = pd.to_datetime(df['Month'])
            recent_df = df[df['Month'] >= pd.Timestamp.now() - pd.DateOffset(months=3)]
            
            if not recent_df.empty:
                category_totals = recent_df.groupby('Category')['Amount'].sum()
                top_categories = category_totals.nlargest(3)
                
                trends += f"Top {sheet} Categories:\n"
                for cat, amount in top_categories.items():
                    trends += f"{cat}: ${amount:,.2f}\n"
                trends += "\n"
        
        return trends

    def _init_budget_file(self):
        """Initialize budget file if it doesn't exist"""
        try:
            if not os.path.exists(self.budget_file):
                with open(self.budget_file, 'w') as f:
                    json.dump({}, f)
        except Exception as e:
            print(f"Error initializing budget file: {str(e)}")

    def set_budget(self, category, amount):
        """Set budget for a category"""
        try:
            # Read existing budgets
            with open(self.budget_file, 'r') as f:
                budgets = json.load(f)
            
            # Update budget for category
            budgets[category] = amount
            
            # Save updated budgets
            with open(self.budget_file, 'w') as f:
                json.dump(budgets, f)
            
            return True
        except Exception as e:
            print(f"Error setting budget: {str(e)}")
            return False

    def get_budget_status(self):
        """Get budget status for all categories"""
        try:
            # Read budgets
            with open(self.budget_file, 'r') as f:
                budgets = json.load(f)
            
            return budgets if budgets else None
            
        except Exception as e:
            print(f"Error getting budget status: {str(e)}")
            return None

    def delete_entry(self, sheet_name, index):
        """Delete an entry from the specified sheet."""
        try:
            # Read all sheets
            all_data = {}
            for sheet in self.sheets:
                df = pd.read_excel(self.file_path, sheet_name=sheet, engine='openpyxl')
                if sheet == sheet_name:
                    # Remove the entry at the specified index
                    df = df.drop(index=index).reset_index(drop=True)
                all_data[sheet] = df
            
            # Write all sheets back in a single operation
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='w') as writer:
                for sheet, data in all_data.items():
                    data.to_excel(writer, sheet_name=sheet, index=False)
            
            return True
        
        except Exception as e:
            print(f"Debug - Delete error: {str(e)}")  # For debugging
            return False 

    def remove_budget(self, category):
        """Remove budget for a specific category"""
        try:
            # Read existing budgets
            with open(self.budget_file, 'r') as f:
                budgets = json.load(f)
            
            # Remove the category if it exists
            if category in budgets:
                del budgets[category]
            
            # Save updated budgets
            with open(self.budget_file, 'w') as f:
                json.dump(budgets, f)
            
            return True
        except Exception as e:
            print(f"Error removing budget: {str(e)}")
            return False 