import pandas as pd
import matplotlib.pyplot as plt

def generate_graphs():
    # --- CONFIGURATION: List of all files and coins to process ---
    tasks = [
        #  The Raw Data File (Coins 5A and 10A)
        {
            "file_name": "group7rawdata.xlsx",
            "start_row": 5,  # Data starts at row 6 (index 5)
            "y_limit": 100,
            "coins": {
                "5A":  {'n': 0, 'h': 3, 't': 4, 'color_h': 'blue', 'color_t': 'orange'},
                "10A": {'n': 6, 'h': 9, 't': 10, 'color_h': 'green', 'color_t': 'red'}
            }
        },
        # The Combined Data File (Combined 5A & 10A)
        {
            "file_name": "raw data and combined data.xlsx",
            "start_row": 5, # Data starts at row 6 (index 5)
            "y_limit": 200,
            "coins": {
                "Combined": {'n': 0, 'h': 3, 't': 4, 'color_h': 'purple', 'color_t': 'brown'}
            }
        }
    ]

    # --- PROCESSING LOOP ---
    for task in tasks:
        file_path = task["file_name"]
        print(f"\n--- Processing File: {file_path} ---")

        current_y_limit = task.get("y_limit", 100)
        
        try:
            # Load the Excel file (header=None to read raw data)
            df = pd.read_excel(file_path, header=None)
        except FileNotFoundError:
            print(f"Error: The file '{file_path}' was not found. Make sure it is in the same folder.")
            continue

        # Process each coin in this file
        for coin_name, config in task["coins"].items():
            print(f"Generating graph for: {coin_name}...")
            
            # Extract specific columns
            # .iloc[rows, [n_col, h_col, t_col]]
            coin_data = df.iloc[task["start_row"]:, [config['n'], config['h'], config['t']]].copy()
            
            # Clean data
            coin_data.columns = ['n', 'Acc_Heads', 'Acc_Tails']
            coin_data = coin_data.apply(pd.to_numeric, errors='coerce').dropna()

            if coin_data.empty:
                print(f"  No valid data for {coin_name}. Skipping.")
                continue

            # Get the very last value in the column to find the final count
            total_heads = int(coin_data['Acc_Heads'].iloc[-1])
            total_tails = int(coin_data['Acc_Tails'].iloc[-1])

            plt.plot(coin_data['n'], coin_data['Acc_Heads'], 
                     label=f'Accumulated Heads ({total_heads})', color=config['color_h'], linewidth=2)
            plt.plot(coin_data['n'], coin_data['Acc_Tails'], 
                     label=f'Accumulated Tails ({total_tails})', color=config['color_t'], linewidth=2)

            # Create Plot
            plt.figure(num=f"Coin {coin_name}", figsize=(10, 6))
            
            plt.plot(coin_data['n'], coin_data['Acc_Heads'], 
                     label='Accumulated Heads', color=config['color_h'], linewidth=2)
            plt.plot(coin_data['n'], coin_data['Acc_Tails'], 
                     label='Accumulated Tails', color=config['color_t'], linewidth=2)

            # Styling

            label_text = f"FINAL RESULT:  Heads = {total_heads}   |   Tails = {total_tails}"

            plt.text(0.5, 0.95, label_text, 
                     transform=plt.gca().transAxes, 
                     ha='center', va='top', 
                     fontsize=12, fontweight='bold', 
                     bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="black", alpha=0.8))
            
            
            plt.title(f'Coin {coin_name}', fontsize=14)
            plt.xlabel('Total Number of Tosses', fontsize=12)
            plt.ylabel('Current Accumulated Count', fontsize=12)
            plt.ylim(0, current_y_limit) 
            plt.legend()
            plt.grid(True, linestyle='--', alpha=0.7)
            
            # Save
            output_file = f"Graph_{coin_name}_Race.png"
            plt.savefig(output_file)
            print(f"  Saved: {output_file}")
            
            plt.show() 

# Run the function

generate_graphs()
