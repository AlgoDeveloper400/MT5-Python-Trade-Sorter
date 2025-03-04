import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
#insert the file path of where your tester report is saved(must be saved as xlsx format for this to work!)
file_path = r"C:\Users\User\Documents\TradingReports\ReportTester-51825733.xlsx"


# Step 1: Load the data without headers (to detect the Deals table)
try:
    print("Loading Excel file...")
    df = pd.read_excel(file_path, sheet_name=0, header=None)  # Load raw data
    print("Excel file loaded successfully.")
except FileNotFoundError:
    print("Error: The file was not found. Please check the file path.")
    exit()
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

# Step 2: Find the first row that contains 'Direction' (start of the Deals table)
deal_table_start = None

for i in range(len(df)):
    row_values = df.iloc[i].astype(str).str.lower()  # Convert row values to lowercase strings
    if 'direction' in row_values.values and 'time' in row_values.values and 'symbol' in row_values.values:
        deal_table_start = i
        break

if deal_table_start is None:
    print("Error: 'Deals' table not found in the Excel file.")
    exit()

# Step 3: Delete all rows above the Deals table and reload with correct headers
df = pd.read_excel(file_path, sheet_name=0, skiprows=deal_table_start)  # Keep only the Deals table

# Step 4: Validate required columns
df.columns = df.columns.str.strip()  # Remove extra spaces
required_columns = ['Direction', 'Time', 'Symbol', 'Type', 'Volume', 'Price', 'Commission', 'Swap', 'Profit', 'Balance', 'Comment']

missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    print(f"Error: Missing columns in the data: {missing_columns}")
    exit()

# Step 5: Remove all empty or unwanted rows
df = df[df['Direction'].notna()].reset_index(drop=True)  # Keep only valid trades

# Step 6: Separate entries and exits
entries = df[df['Direction'] == 'in'].reset_index(drop=True)
exits = df[df['Direction'] == 'out'].reset_index(drop=True)

print(f"Entries found: {len(entries)}, Exits found: {len(exits)}")

if len(entries) != len(exits):
    print("Error: Mismatch between entry and exit trades. Please verify the data.")
    exit()

# Step 7: Combine entry and exit rows into a single row for each trade
combined_data = []
for i in range(len(entries)):
    try:
        combined_data.append({
            "Trade No": i + 1,
            "Entry Time": entries.loc[i, "Time"],
            "Exit Time": exits.loc[i, "Time"],
            "Symbol": entries.loc[i, "Symbol"],
            "Type": entries.loc[i, "Type"],
            "Volume": entries.loc[i, "Volume"],
            "Entry Price": entries.loc[i, "Price"],
            "Exit Price": exits.loc[i, "Price"],
            "Commission": entries.loc[i, "Commission"] + exits.loc[i, "Commission"],
            "Swap": entries.loc[i, "Swap"] + exits.loc[i, "Swap"],
            "Profit": exits.loc[i, "Profit"],
            "Balance": exits.loc[i, "Balance"],
            "Comment": exits.loc[i, "Comment"]
        })
    except Exception as e:
        print(f"Error combining row {i}: {e}")
        exit()

# Step 8: Create DataFrame for combined data
result_df = pd.DataFrame(combined_data)
print("Combined data created successfully.")

# Step 9: Save only the 'deals' table to PDF

#select a file path to save the output pdf file
pdf_file = r"C:\Users\User\Documents\TradingReports\format_trades.pdf"
rows_per_page = 40

try:
    print("Saving data to PDF...")
    with PdfPages(pdf_file) as pdf:
        plt.close('all')  # Remove all graphs, images, and reports

        for start in range(0, len(result_df), rows_per_page):
            end = start + rows_per_page
            subset_df = result_df.iloc[start:end]

            fig, ax = plt.subplots(figsize=(12, 8))
            ax.axis('tight')
            ax.axis('off')

            table = plt.table(
                cellText=subset_df.values,
                colLabels=subset_df.columns,
                cellLoc='center',
                loc='center'
            )
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.auto_set_column_width(col=list(range(len(subset_df.columns))))

            pdf.savefig(fig)
            plt.close(fig)

    print(f"Processed data saved to {pdf_file}.")
except Exception as e:
    print(f"Error saving to PDF: {e}")
