import tkinter as tk
from tkinter import ttk, messagebox, PhotoImage
from finance_tracker import FinanceTracker
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import seaborn as sns
import time

class FinanceTrackerGUI:
    def __init__(self, root):
        """Initialize the Finance Tracker GUI"""
        self.root = root
        self.root.withdraw()  # Hide the main window initially
        
        # Configure color scheme first - before splash screen
        self.setup_styles()
        
        # Now show splash screen
        self.show_splash_screen()
        
        self.root.title("Finance Tracker BY: Augustine Laplace")
        
        # Initialize tracker
        self.tracker = FinanceTracker()
        
        # Create main menu
        self.create_menu()
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create tabs
        self.main_tab = ttk.Frame(self.notebook)
        self.analysis_tab = ttk.Frame(self.notebook)
        self.budget_tab = ttk.Frame(self.notebook)
        
        # Add tabs to notebook
        self.notebook.add(self.main_tab, text='Main')
        self.notebook.add(self.analysis_tab, text='Analysis')
        self.notebook.add(self.budget_tab, text='Budget')
        
        # Setup each tab
        self.setup_main_tab()
        self.setup_analysis_tab()
        self.setup_budget_tab()
        
        # Configure window
        self.root.configure(padx=15, pady=15)
        self.root.resizable(True, True)
        
        # Update window
        self.root.update_idletasks()

    def setup_styles(self):
        """Configure custom styles for the GUI"""
        # Define color palette with modern, harmonious colors
        self.colors = {
            'primary': '#2c3e50',      # Dark blue-gray
            'secondary': '#3498db',    # Bright blue
            'accent': '#e74c3c',       # Soft red
            'background': '#ecf0f1',   # Light gray
            'frame_bg': '#ffffff',     # White
            'text': '#2c3e50',         # Dark blue-gray
            'success': '#2ecc71',      # Emerald green
            'warning': '#f1c40f',      # Sunflower yellow
            'highlight': '#ebf5fb',    # Very light blue
            'border': '#bdc3c7',       # Light gray
            'tab_selected_fg': 'black',  # White
            'menu_bg': '#34495e',      # Darker blue-gray
            'menu_fg': '#ffffff',      # White
            'menu_active': '#3498db'   # Bright blue
        }
        
        # Configure root window
        self.root.configure(bg=self.colors['background'])
        
        # Configure menu colors
        self.root.option_add('*Menu.background', self.colors['menu_bg'])
        self.root.option_add('*Menu.foreground', self.colors['menu_fg'])
        self.root.option_add('*Menu.activeBackground', self.colors['menu_active'])
        self.root.option_add('*Menu.activeForeground', self.colors['menu_fg'])
        self.root.option_add('*Menu.selectColor', self.colors['menu_fg'])
        
        # Configure ttk styles
        style = ttk.Style()
        
        # Configure frame styles
        style.configure('Custom.TFrame',
                       background=self.colors['frame_bg'])
        
        style.configure('Custom.TLabelframe',
                       background=self.colors['frame_bg'],
                       foreground=self.colors['text'])
        
        style.configure('Custom.TLabelframe.Label',
                       foreground=self.colors['primary'],
                       background=self.colors['frame_bg'],
                       font=('Helvetica', 11, 'bold'))
        
        # Configure button style
        style.configure('Custom.TButton',
                       padding=8,
                       font=('Helvetica', 9, 'bold'),
                       borderwidth=1,
                       background=self.colors['secondary'])
        
        style.map('Custom.TButton',
                 background=[('active', self.colors['primary'])],
                 foreground=[('active', self.colors['frame_bg'])])
        
        # Configure label style
        style.configure('Custom.TLabel',
                       background=self.colors['frame_bg'],
                       foreground=self.colors['text'],
                       font=('Helvetica', 10))
        
        # Configure entry style
        style.configure('Custom.TEntry',
                       fieldbackground=self.colors['highlight'],
                       foreground=self.colors['text'],
                       borderwidth=1,
                       relief='solid')
        
        # Configure combobox style
        style.configure('TCombobox',
                       fieldbackground=self.colors['highlight'],
                       background=self.colors['frame_bg'],
                       foreground=self.colors['text'],
                       arrowcolor=self.colors['primary'])
        
        # Configure notebook style
        style.configure('TNotebook',
                       background=self.colors['background'],
                       tabmargins=[2, 5, 2, 0])
        
        style.configure('TNotebook.Tab',
                       background=self.colors['frame_bg'],
                       foreground=self.colors['text'],
                       padding=[15, 5],
                       font=('Helvetica', 10))
        
        style.map('TNotebook.Tab',
                 background=[('selected', self.colors['secondary'])],
                 foreground=[('selected', self.colors['tab_selected_fg'])])
        
        # Configure text widget colors
        text_config = {
            'bg': self.colors['highlight'],
            'fg': self.colors['text'],
            'font': ('Helvetica', 10),
            'relief': 'solid',
            'borderwidth': 1,
            'padx': 5,
            'pady': 5
        }
        
        # Apply text widget configuration
        if hasattr(self, 'summary_text'):
            self.summary_text.configure(**text_config)
        if hasattr(self, 'analysis_text'):
            self.analysis_text.configure(**text_config)
        if hasattr(self, 'budget_text'):
            self.budget_text.configure(**text_config)
        
        # Configure matplotlib style
        plt.style.use('bmh')  # Using a built-in style instead of seaborn
        sns.set_style("whitegrid")
        sns.set_palette([
            self.colors['secondary'],
            self.colors['accent'],
            self.colors['success']
        ])
        
    def create_menu(self):
        """Create the main menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Refresh All", command=self.refresh_all)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # View Menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        
        # Main tab submenu
        main_submenu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="Main", menu=main_submenu)
        main_submenu.add_command(label="Monthly Summary", command=self.update_summary)
        main_submenu.add_command(label="Visualization", command=self.update_visualization)
        main_submenu.add_command(label="Entries View", command=self.update_entries_view)
        
        # Analysis tab submenu
        analysis_submenu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="Analysis", menu=analysis_submenu)
        analysis_submenu.add_command(label="Category Analysis", 
                                   command=lambda: self.notebook.select(1))
        analysis_submenu.add_command(label="Trending Categories", 
                                   command=self.show_trending_categories)
        
        # Budget tab submenu
        budget_submenu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="Budget", menu=budget_submenu)
        budget_submenu.add_command(label="Budget Status", 
                                 command=lambda: self.notebook.select(2))
        budget_submenu.add_command(label="Set Budget", 
                                 command=lambda: self.notebook.select(2))
        
        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        # Instructions submenu
        instructions_menu = tk.Menu(help_menu, tearoff=0)
        help_menu.add_cascade(label="Instructions", menu=instructions_menu)
        instructions_menu.add_command(label="Main Tab", command=self.show_main_instructions)
        instructions_menu.add_command(label="Analysis Tab", command=self.show_analysis_instructions)
        instructions_menu.add_command(label="Budget Tab", command=self.show_budget_instructions)
        
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)

    def refresh_all(self):
        """Refresh all displays"""
        self.update_summary()
        self.update_visualization()
        self.update_entries_view()
        self.update_category_list()
        messagebox.showinfo("Success", "All displays refreshed!")

    def show_about(self):
        """Show about dialog"""
        about_text = """
        Finance Tracker
        Version 1.0
        
        Created by: Augustine Laplace
        
        A simple application to track your
        income, expenses, and savings.
        """
        messagebox.showinfo("About Finance Tracker", about_text)

    def setup_main_tab(self):
        # Create main frames
        self.create_input_frame()
        self.create_summary_frame()
        self.create_visualization_frame()
        self.create_entries_view_frame()
        
    def setup_analysis_tab(self):
        """Create category-based analysis section"""
        # Category Analysis Frame
        analysis_frame = ttk.LabelFrame(self.analysis_tab, text="Category Analysis", padding="10")
        analysis_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Category selection frame
        select_frame = ttk.Frame(analysis_frame)
        select_frame.pack(fill='x', pady=5)
        
        ttk.Label(select_frame, text="Select Category:").pack(side='left', padx=5)
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(select_frame, textvariable=self.category_var, width=30)
        self.category_combo.pack(side='left', padx=5)
        self.update_category_list()
        
        # Analysis text with scrollbar
        text_frame = ttk.Frame(analysis_frame)
        text_frame.pack(fill='both', expand=True, pady=5)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side='right', fill='y')
        
        # Configure analysis text
        self.analysis_text = tk.Text(text_frame, height=15, width=50, yscrollcommand=scrollbar.set)
        self.analysis_text.pack(side='left', fill='both', expand=True, padx=5)
        scrollbar.config(command=self.analysis_text.yview)
        
        # Analysis buttons frame
        btn_frame = ttk.Frame(analysis_frame)
        btn_frame.pack(fill='x', pady=10)
        
        ttk.Button(btn_frame, 
                  text="Show Category Summary",
                  command=self.show_category_analysis,
                  style='Custom.TButton',
                  width=25).pack(side='left', padx=5)
        
        ttk.Button(btn_frame, 
                  text="Show Trending Categories",
                  command=self.show_trending_categories,
                  style='Custom.TButton',
                  width=25).pack(side='left', padx=5)
        
        # Bind tab selection event
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_change)

    def setup_budget_tab(self):
        """Create budget tracking section"""
        budget_frame = ttk.LabelFrame(self.budget_tab, text="Budget Management", padding="10")
        budget_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Budget setting section
        set_budget_frame = ttk.Frame(budget_frame)
        set_budget_frame.pack(fill='x', pady=10)
        
        ttk.Label(set_budget_frame, text="Category:").pack(side='left', padx=5)
        self.budget_category = ttk.Entry(set_budget_frame)
        self.budget_category.pack(side='left', padx=5)
        
        ttk.Label(set_budget_frame, text="Budget Amount:").pack(side='left', padx=5)
        self.budget_amount = ttk.Entry(set_budget_frame)
        self.budget_amount.pack(side='left', padx=5)
        
        ttk.Button(set_budget_frame, 
                  text="Set Budget",
                  command=self.set_budget,
                  style='Custom.TButton').pack(side='left', padx=5)
        
        # Budget status section
        status_frame = ttk.LabelFrame(budget_frame, text="Budget Status", padding="5")
        status_frame.pack(fill='x', pady=5)
        
        # Budget display text
        self.budget_text = tk.Text(status_frame, height=10, width=50)
        self.budget_text.pack(pady=5)
        
        # Status buttons frame
        status_button_frame = ttk.Frame(status_frame)
        status_button_frame.pack(fill='x', pady=5)
        
        # Show Status button
        ttk.Button(status_button_frame, 
                  text="Show Budget Status",
                  command=self.show_budget_status,
                  style='Custom.TButton',
                  width=20).pack(side='left', padx=5)
        
        # Clear Unused button - moved to status frame for better visibility
        ttk.Button(status_button_frame, 
                  text="Clear Unused Budgets",
                  command=self.clear_unused_budgets,
                  style='Custom.TButton',
                  width=20).pack(side='left', padx=5)
        
        # Budget removal section
        remove_frame = ttk.LabelFrame(budget_frame, text="Remove Budgets", padding="5")
        remove_frame.pack(fill='x', pady=5)
        
        # Create listbox for budget selection
        self.budget_listbox = tk.Listbox(remove_frame, selectmode=tk.MULTIPLE, height=5)
        self.budget_listbox.pack(fill='x', pady=5, padx=5)
        
        # Removal buttons frame
        remove_button_frame = ttk.Frame(remove_frame)
        remove_button_frame.pack(fill='x', pady=5)
        
        # Remove Selected button
        ttk.Button(remove_button_frame, 
                  text="Remove Selected Budgets",
                  command=self.remove_selected_budgets,
                  style='Custom.TButton',
                  width=20).pack(side='left', padx=5)
        
        # Initial update of budget list
        self.update_budget_list()

    def create_input_frame(self):
        """Create the input section of the GUI"""
        input_frame = ttk.LabelFrame(self.main_tab, 
                                   text="Add New Entry", 
                                   padding="10",
                                   style='Custom.TLabelframe')
        input_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        
        # Entry Type Selection
        ttk.Label(input_frame, 
                 text="Entry Type:",
                 style='Custom.TLabel').grid(row=0, column=0, padx=5, pady=5)
        self.entry_type = ttk.Combobox(input_frame, 
                                      values=self.tracker.sheets,
                                      style='Custom.TCombobox')
        self.entry_type.grid(row=0, column=1, padx=5, pady=5)
        self.entry_type.set(self.tracker.sheets[0])
        
        # Bind entry type selection to update categories
        self.entry_type.bind('<<ComboboxSelected>>', self.update_entry_categories)
        
        # Amount Input
        ttk.Label(input_frame, text="Amount:").grid(row=1, column=0, padx=5, pady=5)
        self.amount = ttk.Entry(input_frame)
        self.amount.grid(row=1, column=1, padx=5, pady=5)
        
        # Category Input with suggestions
        ttk.Label(input_frame, text="Category:").grid(row=2, column=0, padx=5, pady=5)
        self.category = ttk.Combobox(input_frame)
        self.category.grid(row=2, column=1, padx=5, pady=5)
        
        # Notes Input
        ttk.Label(input_frame, text="Notes:").grid(row=3, column=0, padx=5, pady=5)
        self.notes = ttk.Entry(input_frame)
        self.notes.grid(row=3, column=1, padx=5, pady=5)
        
        # Submit Button
        submit_btn = ttk.Button(input_frame, 
                               text="Add Entry", 
                               command=self.add_entry,
                               style='Custom.TButton')
        submit_btn.grid(row=4, column=0, columnspan=2, pady=10)
        
        # Initialize categories for default entry type
        self.update_entry_categories()

    def update_entry_categories(self, event=None):
        """Update category suggestions based on selected entry type"""
        try:
            entry_type = self.entry_type.get()
            if not entry_type:
                return
            
            # Read the selected sheet
            df = pd.read_excel(self.tracker.file_path, sheet_name=entry_type, engine='openpyxl')
            
            # Get unique categories from the selected sheet
            categories = df['Category'].dropna().unique().tolist()
            
            # Sort categories and update combobox
            categories.sort()
            self.category['values'] = categories
            
            # Clear current category
            self.category.set('')
            
        except Exception as e:
            print(f"Debug - Category update error: {str(e)}")  # For debugging
            # If there's an error, just clear the categories
            self.category['values'] = []
            self.category.set('')

    def update_category_list(self):
        """Update the category dropdown with existing categories"""
        categories = self.tracker.get_all_categories()
        self.category_combo['values'] = list(categories)
        self.category['values'] = list(categories)
    
    def show_category_analysis(self):
        """Display analysis for selected category"""
        category = self.category_var.get()
        if not category:
            messagebox.showwarning("Warning", "Please select a category")
            return
            
        analysis = self.tracker.analyze_category(category)
        self.analysis_text.delete(1.0, tk.END)
        self.analysis_text.insert(tk.END, analysis)
    
    def show_trending_categories(self):
        """Show trending categories based on recent transactions"""
        trends = self.tracker.get_trending_categories()
        self.analysis_text.delete(1.0, tk.END)
        self.analysis_text.insert(tk.END, trends)
    
    def show_budget_status(self):
        """Display current budget status"""
        try:
            # Get budget status from tracker
            status = self.tracker.get_budget_status()
            
            if not status:
                self.budget_text.delete(1.0, tk.END)
                self.budget_text.insert(tk.END, "No budgets set yet. Please set a budget first.")
                return
            
            # Clear previous text
            self.budget_text.delete(1.0, tk.END)
            
            # Get current month's expenses
            current_month = datetime.now().strftime('%Y-%m')
            
            try:
                df = pd.read_excel(self.tracker.file_path, sheet_name='Expenses', engine='openpyxl')
                
                # Ensure Amount column is numeric
                df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
                
                # Convert dates to consistent format
                df['Month'] = pd.to_datetime(df['Month']).dt.strftime('%Y-%m')
                
                # Filter for current month
                current_expenses = df[df['Month'] == current_month]
                
                # Format and display budget status
                status_text = f"Budget Status for {current_month}:\n\n"
                
                for category, budget in status.items():
                    # Calculate spent amount for this category
                    category_expenses = current_expenses[
                        current_expenses['Category'].str.lower().str.strip() == category.lower().strip()
                    ]
                    
                    spent = float(category_expenses['Amount'].sum()) if not category_expenses.empty else 0.0
                    remaining = float(budget) - spent
                    
                    status_text += f"Category: {category}\n"
                    status_text += f"Budget: ${float(budget):,.2f}\n"
                    status_text += f"Spent: ${spent:,.2f}\n"
                    status_text += f"Remaining: ${remaining:,.2f}\n"
                    
                    if remaining < 0:
                        status_text += "WARNING: Over budget!\n"
                    
                    status_text += "-" * 30 + "\n"
                
                self.budget_text.insert(tk.END, status_text)
                
            except Exception as e:
                raise
            
        except Exception as e:
            self.show_error('operation', 'budget_status')

    def set_budget(self):
        """Set budget for a category"""
        try:
            # Validate category
            category = self.budget_category.get().strip()
            if not category:
                self.show_error('input', 'category')
                return
            
            # Validate amount
            try:
                amount = float(self.budget_amount.get())
                if amount <= 0:
                    raise ValueError("Amount must be positive")
            except ValueError:
                self.show_error('input', 'budget')
                return
            
            # Set the budget
            if self.tracker.set_budget(category, amount):
                # Clear inputs
                self.budget_category.delete(0, tk.END)
                self.budget_amount.delete(0, tk.END)
                
                # Show success message and update displays
                messagebox.showinfo("Success", f"Budget of ${amount:,.2f} set for {category}")
                self.update_budget_list()
                self.show_budget_status()
            
        except Exception as e:
            self.show_error('operation', 'update')
    
    def create_summary_frame(self):
        """Create the summary section of the GUI"""
        summary_frame = ttk.LabelFrame(self.main_tab, text="Financial Summary", padding="10")
        summary_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        self.summary_text = tk.Text(summary_frame, height=5, width=40)
        self.summary_text.grid(row=0, column=0, padx=5, pady=5)
        
        refresh_btn = ttk.Button(summary_frame, 
                                text="Refresh Summary", 
                                command=self.update_summary,
                                style='Custom.TButton')
        refresh_btn.grid(row=1, column=0, pady=5)
        
    def create_visualization_frame(self):
        """Create the visualization section of the GUI"""
        viz_frame = ttk.LabelFrame(self.main_tab, 
                                 text="Visualization", 
                                 padding="10",
                                 style='Custom.TLabelframe')
        viz_frame.grid(row=0, column=1, rowspan=2, padx=10, pady=5, sticky="nsew")
        
        # Create figure with custom styling
        self.fig, self.ax = plt.subplots(figsize=(6, 4))
        self.fig.patch.set_facecolor(self.colors['background'])
        self.ax.set_facecolor(self.colors['background'])
        
        # Set custom color palette
        custom_palette = [self.colors['secondary'], 
                         self.colors['accent'], 
                         self.colors['success']]
        sns.set_palette(custom_palette)
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=viz_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        refresh_btn = ttk.Button(viz_frame, 
                               text="Refresh Chart",
                               command=self.update_visualization,
                               style='Custom.TButton')
        refresh_btn.pack(pady=5)
        
    def add_entry(self):
        """Handle adding new entries"""
        try:
            # Validate amount
            try:
                amount = float(self.amount.get())
                if amount <= 0:
                    raise ValueError("Amount must be positive")
            except ValueError:
                self.show_error('input', 'amount')
                return
            
            # Validate category
            category = self.category.get().strip()
            if not category:
                self.show_error('input', 'category')
                return
            
            entry_type = self.entry_type.get()
            notes = self.notes.get()
            
            if self.tracker.add_entry(entry_type, amount, category, notes):
                # Clear inputs
                self.amount.delete(0, tk.END)
                self.category.delete(0, tk.END)
                self.notes.delete(0, tk.END)
                
                # Update display
                self.update_summary()
                self.update_visualization()
                self.update_entries_view()
                
                messagebox.showinfo("Success", "Entry added successfully!")
        except Exception as e:
            self.show_error('operation', 'add')
    
    def update_summary(self):
        """Update the summary text"""
        summary = self.tracker.get_monthly_summary()
        
        summary_text = "Current Month Summary:\n\n"
        for category, amount in summary.items():
            summary_text += f"{category}: ${amount:,.2f}\n"
            
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.insert(tk.END, summary_text)
    
    def update_visualization(self):
        """Update the visualization chart with custom styling"""
        self.ax.clear()
        
        data = []
        for sheet in self.tracker.sheets:
            df = self.tracker.get_sheet_data(sheet)
            if not df.empty:
                data.append(df)
        
        if data:
            plot_data = pd.concat(data)
            
            # Get unique types in the data
            unique_types = plot_data['Type'].unique()
            
            # Create color mapping based on actual types present
            color_mapping = {
                'Income': self.colors['success'],
                'Expenses': self.colors['accent'],
                'Savings': self.colors['secondary']
            }
            
            # Filter palette to only include colors for present types
            palette = [color_mapping[type_] for type_ in unique_types]
            
            sns.barplot(data=plot_data, 
                       x='Month', 
                       y='Amount', 
                       hue='Type', 
                       ax=self.ax,
                       palette=palette)
            
            self.ax.set_title('Monthly Financial Summary', 
                            color=self.colors['primary'],
                            fontsize=12,
                            pad=15)
            
            self.ax.tick_params(axis='both', 
                              colors=self.colors['text'])
            
            self.ax.set_xlabel('Month', 
                             color=self.colors['text'],
                             fontsize=10)
            self.ax.set_ylabel('Amount', 
                             color=self.colors['text'],
                             fontsize=10)
            
            # Rotate x-axis labels
            plt.xticks(rotation=45)
            
            # Add grid with custom style
            self.ax.grid(True, 
                        linestyle='--', 
                        alpha=0.7, 
                        color=self.colors['text'])
            
            plt.tight_layout()
            self.canvas.draw()

    def create_entries_view_frame(self):
        """Create a frame to view and delete entries"""
        entries_frame = ttk.LabelFrame(self.main_tab, 
                                    text="View/Delete Entries", 
                                    padding="10",
                                    style='Custom.TLabelframe')
        entries_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        
        # Create a frame for the sheet selection
        select_frame = ttk.Frame(entries_frame)
        select_frame.pack(fill='x', pady=5)
        
        ttk.Label(select_frame, text="Select Sheet:").pack(side='left', padx=5)
        self.sheet_var = tk.StringVar()
        sheet_combo = ttk.Combobox(select_frame, 
                                  textvariable=self.sheet_var,
                                  values=self.tracker.sheets)
        sheet_combo.pack(side='left', padx=5)
        sheet_combo.set(self.tracker.sheets[0])
        
        # Create Treeview for entries
        columns = ('Date', 'Amount', 'Category', 'Notes')
        self.entries_tree = ttk.Treeview(entries_frame, columns=columns, show='headings')
        
        # Configure columns
        for col in columns:
            self.entries_tree.heading(col, text=col)
            self.entries_tree.column(col, width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(entries_frame, orient='vertical', command=self.entries_tree.yview)
        self.entries_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack the Treeview and scrollbar
        self.entries_tree.pack(side='left', fill='both', expand=True, pady=5)
        scrollbar.pack(side='right', fill='y')
        
        # Create button frame
        button_frame = ttk.Frame(entries_frame)
        button_frame.pack(pady=5)
        
        # Add buttons
        view_btn = ttk.Button(button_frame,
                             text="View Entry",
                             command=self.view_selected_entry,
                             style='Custom.TButton')
        view_btn.pack(side='left', padx=5)
        
        edit_btn = ttk.Button(button_frame,
                             text="Edit Entry",
                             command=self.edit_selected_entry,
                             style='Custom.TButton')
        edit_btn.pack(side='left', padx=5)
        
        delete_btn = ttk.Button(button_frame,
                               text="Delete Entry",
                               command=self.delete_selected_entry,
                               style='Custom.TButton')
        delete_btn.pack(side='left', padx=5)
        
        # Bind the sheet selection to update the view
        sheet_combo.bind('<<ComboboxSelected>>', self.update_entries_view)
        
        # Initial load of entries
        self.update_entries_view()

    def view_selected_entry(self):
        """View details of the selected entry"""
        selected = self.entries_tree.selection()
        if not selected:
            self.show_error('input', 'selection')
            return
        
        try:
            sheet = self.sheet_var.get()
            index = int(selected[0])
            
            # Get the entry details
            df = pd.read_excel(self.tracker.file_path, sheet_name=sheet, engine='openpyxl')
            entry = df.iloc[index]
            
            # Format the details
            details = f"""Entry Details:

Type: {sheet}
Date: {entry['Month']}
Amount: ${entry['Amount']:,.2f}
Category: {entry['Category']}
Notes: {entry['Notes']}"""
            
            messagebox.showinfo("Entry Details", details)
            
        except Exception as e:
            self.show_error('operation', 'view')

    def edit_selected_entry(self):
        """Edit the selected entry"""
        selected = self.entries_tree.selection()
        if not selected:
            self.show_error('input', 'selection')
            return
        
        try:
            sheet = self.sheet_var.get()
            index = int(selected[0])
            
            # Get the entry details
            df = pd.read_excel(self.tracker.file_path, sheet_name=sheet, engine='openpyxl')
            entry = df.iloc[index]
            
            # Create edit dialog
            edit_window = tk.Toplevel(self.root)
            edit_window.title("Edit Entry")
            edit_window.geometry("300x250")
            edit_window.resizable(False, False)
            
            # Add padding
            edit_window.configure(padx=10, pady=10)
            
            # Entry fields
            ttk.Label(edit_window, text="Amount:").pack(pady=5)
            amount_entry = ttk.Entry(edit_window)
            amount_entry.insert(0, str(entry['Amount']))
            amount_entry.pack(pady=5)
            
            ttk.Label(edit_window, text="Category:").pack(pady=5)
            category_entry = ttk.Entry(edit_window)
            category_entry.insert(0, entry['Category'])
            category_entry.pack(pady=5)
            
            ttk.Label(edit_window, text="Notes:").pack(pady=5)
            notes_entry = ttk.Entry(edit_window)
            notes_entry.insert(0, str(entry['Notes']))  # Convert to string
            notes_entry.pack(pady=5)
            
            def save_changes():
                try:
                    # Validate amount
                    try:
                        new_amount = float(amount_entry.get())
                        if new_amount <= 0:
                            raise ValueError("Amount must be positive")
                    except ValueError:
                        self.show_error('input', 'amount')
                        return
                    
                    # Validate category
                    new_category = category_entry.get().strip()
                    if not new_category:
                        self.show_error('input', 'category')
                        return
                    
                    # Get all sheets data
                    all_data = {}
                    for s in self.tracker.sheets:
                        if s == sheet:
                            # Update the current sheet
                            temp_df = df.copy()
                            temp_df.at[index, 'Amount'] = new_amount
                            temp_df.at[index, 'Category'] = new_category
                            temp_df.at[index, 'Notes'] = notes_entry.get()
                            all_data[s] = temp_df
                        else:
                            # Preserve other sheets
                            all_data[s] = pd.read_excel(self.tracker.file_path, 
                                                      sheet_name=s, 
                                                      engine='openpyxl')
                    
                    # Save all changes in a single operation
                    with pd.ExcelWriter(self.tracker.file_path, engine='openpyxl', mode='w') as writer:
                        for sheet_name, data in all_data.items():
                            data.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Update display
                    self.update_entries_view()
                    self.update_summary()
                    self.update_visualization()
                    
                    messagebox.showinfo("Success", "Entry updated successfully!")
                    edit_window.destroy()
                    
                except Exception as e:
                    self.show_error('operation', 'update')
                    print(f"Debug - Error details: {str(e)}")  # For debugging
            
            # Save button
            ttk.Button(edit_window,
                      text="Save Changes",
                      command=save_changes,
                      style='Custom.TButton').pack(pady=20)
            
        except Exception as e:
            self.show_error('operation', 'edit')
            print(f"Debug - Error details: {str(e)}")  # For debugging

    def delete_selected_entry(self):
        """Delete the selected entry"""
        selected = self.entries_tree.selection()
        if not selected:
            self.show_error('input', 'selection')
            return
        
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this entry?"):
            try:
                sheet = self.sheet_var.get()
                index = int(selected[0])
                
                # Attempt to delete the entry
                success = self.tracker.delete_entry(sheet, index)
                
                if success:
                    # Update all displays
                    self.update_entries_view()
                    self.update_summary()
                    self.update_visualization()
                    self.update_category_list()
                    messagebox.showinfo("Success", "Entry deleted successfully!")
                else:
                    self.show_error('operation', 'delete')
                
            except Exception as e:
                print(f"Debug - GUI delete error: {str(e)}")  # For debugging
                self.show_error('operation', 'delete')

    def resize_window(self):
        """Resize window to fit current tab content"""
        current_tab = self.notebook.select()
        if current_tab:
            tab = self.notebook.index(current_tab)
            if tab == 0:  # Main tab
                width = 1200  # Width to accommodate visualization
                height = 800
            elif tab == 1:  # Analysis tab
                width = 600
                height = 600
            else:  # Budget tab
                width = 600
                height = 500
            
            # Get screen dimensions
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            # Calculate center position
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            
            # Set window size and position
            self.root.geometry(f"{width}x{height}+{x}+{y}")

    def on_tab_change(self, event):
        """Handle tab change event"""
        current_tab = self.notebook.select()
        tab_id = self.notebook.index(current_tab)
        
        if tab_id == 1:  # Analysis tab
            # Adjust window size for analysis tab
            width = 600
            height = 500
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            self.root.geometry(f"{width}x{height}+{x}+{y}")
        else:
            # Reset to default size for other tabs
            width = 1200
            height = 800
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            self.root.geometry(f"{width}x{height}+{x}+{y}")

    def show_main_instructions(self):
        """Show instructions for the Main tab"""
        instructions = """
Main Tab Instructions:

1. Add New Entry Section:
   • Select entry type (Income/Expenses/Savings)
   • Enter the amount
   • Choose or type a category
   • Add optional notes
   • Click 'Add Entry' to save

2. Financial Summary:
   • Shows total amounts for each category
   • Click 'Refresh Summary' to update

3. Visualization:
   • Bar chart showing monthly trends
   • Different colors for Income/Expenses/Savings
   • Click 'Refresh Chart' to update

4. View/Delete Entries:
   • Select a sheet to view entries
   • Click on an entry to select it
   • Use 'Delete Selected' to remove entries
"""
        messagebox.showinfo("Main Tab Help", instructions)

    def show_analysis_instructions(self):
        """Show instructions for the Analysis tab"""
        instructions = """
Analysis Tab Instructions:

1. Category Analysis:
   • Select a category from the dropdown
   • Click 'Show Category Summary' to view:
     - Total amount per category
     - Average amount per category
     - Breakdown by entry type

2. Trending Categories:
   • Click 'Show Trending Categories' to view:
     - Top categories by amount
     - Last 3 months of data
     - Separated by entry type

3. Tips:
   • Use this tab to track spending patterns
   • Identify major expense categories
   • Monitor income sources
"""
        messagebox.showinfo("Analysis Tab Help", instructions)

    def show_budget_instructions(self):
        """Show instructions for the Budget tab"""
        instructions = """
Budget Tab Instructions:

1. Set Budget:
   • Enter a category name
   • Enter the budget amount
   • Click 'Set Budget' to save
   • You can set multiple category budgets

2. Budget Status:
   • Click 'Show Budget Status' to view:
     - Current spending vs budget
     - Remaining budget amounts
     - Budget alerts if overspent

3. Tips:
   • Set realistic budget amounts
   • Regular monitoring helps stay on track
   • Update budgets as needed
"""
        messagebox.showinfo("Budget Tab Help", instructions)

    def show_error(self, error_type, details):
        """Display customized error messages"""
        error_messages = {
            'input': {
                'title': "Input Error",
                'messages': {
                    'amount': "Please enter a valid number for the amount.",
                    'category': "Category cannot be empty.",
                    'type': "Please select a valid entry type.",
                    'budget': "Please enter a valid budget amount (positive number).",
                    'selection': "Please select an entry first.",
                    'default': "Please check your input and try again."
                }
            },
            'file': {
                'title': "File Error",
                'messages': {
                    'not_found': "The finance tracker file could not be found.",
                    'access': "Unable to access the file. Please check if it's open in another program.",
                    'corrupt': "The file appears to be corrupted. A backup may be needed.",
                    'default': "An error occurred while accessing the file."
                }
            },
            'operation': {
                'title': "Operation Error",
                'messages': {
                    'delete': "Unable to delete the selected entry.",
                    'add': "Failed to add the new entry.",
                    'update': "Failed to update the entry.",
                    'view': "Unable to view the entry details.",
                    'edit': "Unable to edit the entry.",
                    'refresh': "Unable to refresh the display.",
                    'default': "The operation could not be completed."
                }
            },
            'default': {
                'title': "Error",
                'messages': {
                    'default': "An unexpected error occurred. Please try again."
                }
            }
        }

        # Get error category and specific message
        error_info = error_messages.get(error_type, error_messages['default'])
        title = error_info['title']
        message = error_info['messages'].get(details, error_info['messages']['default'])
        
        # Add technical details if available
        if isinstance(details, Exception):
            message += f"\n\nTechnical details:\n{str(details)}"
        
        messagebox.showerror(title, message)

    def update_entries_view(self, event=None):
        """Update the entries view when sheet is selected"""
        sheet = self.sheet_var.get()
        if not sheet:
            return
        
        try:
            # Clear existing items
            for item in self.entries_tree.get_children():
                self.entries_tree.delete(item)
            
            # Read the sheet using the tracker's file path
            df = pd.read_excel(self.tracker.file_path, sheet_name=sheet, engine='openpyxl')
            
            # Add entries to treeview
            for idx, row in df.iterrows():
                self.entries_tree.insert('', 'end', iid=str(idx), values=(
                    row['Month'],
                    f"${row['Amount']:,.2f}",
                    row['Category'],
                    row['Notes']
                ))
        except FileNotFoundError:
            self.show_error('file', 'not_found')
        except PermissionError:
            self.show_error('file', 'access')
        except Exception as e:
            self.show_error('file', 'default')

    def clear_unused_budgets(self):
        """Remove budgets for categories with no recent expenses"""
        try:
            # Get current budgets
            status = self.tracker.get_budget_status()
            if not status:
                messagebox.showinfo("Info", "No budgets to clear.")
                return
            
            # Get all expense categories from the last 3 months
            df = pd.read_excel(self.tracker.file_path, sheet_name='Expenses', engine='openpyxl')
            df['Month'] = pd.to_datetime(df['Month'])
            recent_df = df[df['Month'] >= pd.Timestamp.now() - pd.DateOffset(months=3)]
            active_categories = set(recent_df['Category'].str.lower().str.strip())
            
            # Find unused budget categories
            unused_categories = []
            for category in status.keys():
                if category.lower().strip() not in active_categories:
                    unused_categories.append(category)
            
            if not unused_categories:
                messagebox.showinfo("Info", "No unused budgets found.")
                return
            
            # Confirm with user
            message = "The following unused budgets will be removed:\n\n"
            message += "\n".join(unused_categories)
            message += "\n\nDo you want to continue?"
            
            if messagebox.askyesno("Confirm Clear", message):
                # Remove unused budgets
                for category in unused_categories:
                    self.tracker.remove_budget(category)
                
                messagebox.showinfo("Success", f"Removed {len(unused_categories)} unused budgets.")
                self.show_budget_status()
        
        except Exception as e:
            self.show_error('operation', 'clear_budgets')

    def update_budget_list(self):
        """Update the list of budgets in the listbox"""
        try:
            # Clear current list
            self.budget_listbox.delete(0, tk.END)
            
            # Get current budgets
            status = self.tracker.get_budget_status()
            if status:
                # Add each budget to the list with its amount
                for category, amount in status.items():
                    self.budget_listbox.insert(tk.END, f"{category} (${float(amount):,.2f})")
        except Exception as e:
            self.show_error('operation', 'update_list')

    def remove_selected_budgets(self):
        """Remove selected budgets from the list"""
        try:
            # Get selected indices
            selected_indices = self.budget_listbox.curselection()
            if not selected_indices:
                messagebox.showinfo("Info", "Please select budgets to remove.")
                return
            
            # Get all budget items
            selected_items = [self.budget_listbox.get(idx) for idx in selected_indices]
            
            # Extract category names (remove the amount part)
            selected_categories = [item.split(" ($")[0] for item in selected_items]
            
            # Confirm with user
            message = "The following budgets will be removed:\n\n"
            message += "\n".join(selected_items)
            message += "\n\nDo you want to continue?"
            
            if messagebox.askyesno("Confirm Remove", message):
                # Remove selected budgets
                for category in selected_categories:
                    self.tracker.remove_budget(category)
                
                messagebox.showinfo("Success", f"Removed {len(selected_categories)} budget(s).")
                
                # Update displays
                self.update_budget_list()
                self.show_budget_status()
        
        except Exception as e:
            self.show_error('operation', 'remove_budgets')

    def show_splash_screen(self):
        """Show splash screen with loading progress"""
        # Create splash screen window
        splash = tk.Toplevel(self.root)
        splash.title("Loading Finance Tracker")
        
        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Set splash window size and position
        width = 400
        height = 200
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        splash.geometry(f"{width}x{height}+{x}+{y}")
        
        # Configure splash window with colors from setup_styles
        splash.configure(bg=self.colors['background'])
        splash.overrideredirect(True)  # Remove window decorations
        
        # Add application title
        title_label = tk.Label(splash,
                              text="Finance Tracker",
                              font=('Helvetica', 16, 'bold'),
                              bg=self.colors['background'],
                              fg=self.colors['primary'])
        title_label.pack(pady=20)
        
        # Add loading message
        loading_label = tk.Label(splash,
                                text="Loading...",
                                font=('Helvetica', 10),
                                bg=self.colors['background'],
                                fg=self.colors['text'])
        loading_label.pack(pady=5)
        
        # Create progress bar
        progress = ttk.Progressbar(splash,
                                  length=300,
                                  mode='determinate')
        progress.pack(pady=20)
        
        def update_progress():
            """Update progress bar and loading message"""
            steps = ['Initializing application...',
                    'Loading data...',
                    'Setting up interface...',
                    'Preparing visualizations...',
                    'Almost ready...']
            
            for i, message in enumerate(steps, 1):
                progress['value'] = i * 20
                loading_label.config(text=message)
                splash.update()
                time.sleep(0.3)  # Reduced delay for smoother experience
            
            # Final update
            progress['value'] = 100
            loading_label.config(text="Ready!")
            splash.update()
            time.sleep(0.2)  # Reduced final delay
            
            # Close splash screen and show main window
            splash.destroy()
            self.root.deiconify()
        
        # Start progress update
        splash.after(100, update_progress)

def main():
    try:
        root = tk.Tk()
        
        # Set initial window size and position
        width = 1200
        height = 800
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        root.geometry(f"{width}x{height}+{x}+{y}")
        
        # Create the application
        app = FinanceTrackerGUI(root)
        
        # Start the main event loop
        root.mainloop()
        
    except Exception as e:
        print(f"Error starting application: {str(e)}")
        raise

if __name__ == "__main__":
    main() 