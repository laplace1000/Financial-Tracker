from finance_tracker import FinanceTracker

def main():
    # Create a new finance tracker instance
    tracker = FinanceTracker()
    
    # Add some sample entries
    tracker.add_entry('Income', 5000, 'Salary', 'Monthly salary')
    tracker.add_entry('Expenses', 1500, 'Rent', 'Monthly rent')
    tracker.add_entry('Expenses', 500, 'Groceries', 'Monthly groceries')
    tracker.add_entry('Savings', 1000, 'Emergency Fund', 'Monthly savings')
    
    # Get monthly summary
    summary = tracker.get_monthly_summary()
    print("\nMonthly Summary:")
    for category, amount in summary.items():
        print(f"{category}: ${amount:,.2f}")
    
    # Visualize the data
    tracker.visualize_monthly_data()

if __name__ == "__main__":
    main() 