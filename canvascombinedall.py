import pandas as pd
import matplotlib.pyplot as plt
import os
import re

def analyze_global_fairness_excel():
    # --- CONFIGURATION ---
    excel_file = "2BSCS-A _ Tossed Coin Raw Data.xlsx"
    
    print(f"Working Directory: {os.getcwd()}")
    print(f"Reading {excel_file} to aggregate GLOBAL data...")

    try:
        # Read ALL sheets at once
        all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
        print("File loaded successfully. Processing sheets...")
    except FileNotFoundError:
        print(f"Error: '{excel_file}' not found.")
        return

    # Configuration: Group Number -> [(Coin Name, Heads_Col, Tails_Col)]
    # (Indices are 0-based)
    file_configs = {
        1: [("1 Peso", 3, 2), ("2 Peso", 6, 5)],
        2: [("1B", 3, 4), ("5A", 11, 12)],
        3: [("1B", 1, 2), ("10A", 9, 10)],
        4: [("5A", 3, 4), ("5B", 8, 9)],
        5: [("1A", 3, 4), ("1B", 9, 10)],
        6: [("5B", 3, 4), ("20", 8, 9)],
        7: [("5A", 4, 5), ("10A", 12, 13)],
        8: [("1A", 3, 4), ("10B", 9, 10)],
        9: [("5B", 3, 4), ("1B", 7, 8), ("20", 11, 12)],
        10: [("5B", 3, 4), ("10B", 8, 9)],
        11: [("1A", 4, 5), ("10B", 10, 11)],
        12: [("5B", 2, 4), ("5A", 8, 10)],
        13: [("1A", 1, 2), ("10A", 5, 6)],
        14: [("1A", 1, 2), ("20", 3, 4)],
        15: [("1B", 1, 2), ("5B", 9, 10)]
    }

    grand_total_heads = 0
    grand_total_tails = 0

    # --- ITERATE THROUGH GROUPS 1-15 ---
    for i in range(1, 16):
        group_num = i
        
        # Find the correct sheet name (e.g., "Group 1" or "GROUP 1")
        target_sheet = None
        for sheet_name in all_sheets.keys():
            if re.search(r'GROUP[\s_-]*' + str(group_num) + r'(?![0-9])', sheet_name, re.IGNORECASE):
                target_sheet = sheet_name
                break
        
        if not target_sheet:
            print(f"  Warning: Could not find sheet for Group {group_num}")
            continue

        # Extract Data
        df = all_sheets[target_sheet]
        configs = file_configs.get(group_num, [])

        for _, h_idx, t_idx in configs:
            try:
                # Read columns (skip first 2 header rows)
                col_h = pd.to_numeric(df.iloc[2:, h_idx], errors='coerce').fillna(0)
                col_t = pd.to_numeric(df.iloc[2:, t_idx], errors='coerce').fillna(0)
                
                # Logic: If numbers are big (>10), it's cumulative -> Take Max
                # If numbers are small (0s and 1s), it's tally -> Take Sum
                if col_h.max() > 10 or col_t.max() > 10:
                    grand_total_heads += col_h.max()
                    grand_total_tails += col_t.max()
                else:
                    grand_total_heads += col_h.sum()
                    grand_total_tails += col_t.sum()
            except Exception as e:
                print(f"  Error extracting data in Group {group_num}: {e}")

    # --- CALCULATE FINAL STATS ---
    total_flips = grand_total_heads + grand_total_tails
    if total_flips == 0:
        print("No data extracted.")
        return

    pct_heads = (grand_total_heads / total_flips) * 100
    pct_tails = (grand_total_tails / total_flips) * 100

    print(f"\n--- GLOBAL RESULTS ---")
    print(f"Total Flips: {int(total_flips)}")
    print(f"Total Heads: {int(grand_total_heads)} ({pct_heads:.2f}%)")
    print(f"Total Tails: {int(grand_total_tails)} ({pct_tails:.2f}%)")

    # --- GENERATE GRAPH ---
    labels = ['Heads', 'Tails']
    percentages = [pct_heads, pct_tails]
    colors = ['#1f77b4', '#ff7f0e'] # Blue, Orange

    plt.figure(figsize=(8, 6))
    bars = plt.bar(labels, percentages, color=colors, width=0.6)

    # Add 50% Reference Line
    plt.axhline(y=50, color='red', linestyle='--', linewidth=2, label='Fairness Line (50%)')

    plt.title('Coin Toss Fairness (All Groups and COin Class Combined)', fontsize=14, fontweight='bold')
    plt.ylabel('Percentage (%)', fontsize=12)
    plt.ylim(0, 100) # Zoom in to see the difference clearly
    plt.legend()
    plt.grid(axis='y', linestyle='--', alpha=0.5)

    # Add Labels on Bars
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, height + 1,
                 f'{height:.2f}%', ha='center', va='bottom', fontsize=12, fontweight='bold')

    # Add Counts inside Bars
    plt.text(bars[0].get_x() + bars[0].get_width()/2, bars[0].get_height()/2,
             f"{int(grand_total_heads)}", ha='center', color='white', fontweight='bold', fontsize=14)
    plt.text(bars[1].get_x() + bars[1].get_width()/2, bars[1].get_height()/2,
             f"{int(grand_total_tails)}", ha='center', color='white', fontweight='bold', fontsize=14)

    # Save
    output_file = "Global_Fairness_Graph.png"
    plt.tight_layout()
    plt.savefig(output_file)
    print(f"\nGraph saved as: {output_file}")
    print("Graph window opened...")
    plt.show()

if __name__ == "__main__":
    analyze_global_fairness_excel()