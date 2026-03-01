import pandas as pd
import math
import argparse
import sys
import re

def extract_dims_from_profile(profile_str):
    # Extracts Thickness and Width from formats like PL10*135 or PL 10*135
    match = re.search(r'PL\s*(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)', str(profile_str).upper())
    if match:
        return float(match.group(1)), float(match.group(2))
    return None, None

def process_procurement(input_file, output_file):
    try:
        df = pd.read_csv(input_file) if input_file.endswith('.csv') else pd.read_excel(input_file)
    except Exception as e:
        print(f"Error reading file: {e}"); sys.exit(1)

    # Convert numeric columns safely
    for col in ['Length', 'Width', 'Quantity']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Normalization (mm to m)
    avg_length = df['Length'].mean()
    if avg_length > 100:
        df['Length'] /= 1000
        df['Width'] /= 1000

    # Plate Logic with improved Parsing
    plate_summary = []
    # Identify plates by name starting with PL
    is_plate = df['Profile'].astype(str).str.upper().str.startswith('PL')
    plates_df = df[is_plate].copy()
    profiles_df = df[~is_plate].copy()

    if not plates_df.empty:
        for idx, row in plates_df.iterrows():
            t_ext, w_ext = extract_dims_from_profile(row['Profile'])
            # Use extracted width if the Excel column is 0
            actual_width = (w_ext / 1000) if (row['Width'] == 0 and w_ext) else row['Width']
            actual_thickness = t_ext if t_ext else 0
            
            plates_df.at[idx, 'Thickness'] = actual_thickness
            plates_df.at[idx, 'Actual_Width'] = actual_width

        plate_groups = plates_df.groupby(['Thickness', 'Grade'])
        for (thickness, grade), group in plate_groups:
            total_area = (group['Length'] * group['Actual_Width'] * group['Quantity']).sum()
            # Waste set to 2% (1.02) to be safe for a technical office
            plates_needed = math.ceil((total_area * 1.02) / 18)
            plate_summary.append({
                'Thickness': thickness, 'Grade': grade, 
                'Total Area (sqm)': round(total_area, 3),
                'Standard Plates (1.5x12m) (2% waste)': plates_needed
            })

    # Profile Logic (Beams/Angles)
    profile_summary = []
    if not profiles_df.empty:
        profile_groups = profiles_df.groupby(['Profile', 'Grade'])
        for (profile, grade), group in profile_groups:
            total_length = (group['Length'] * group['Quantity']).sum()
            # Waste set to 3% (1.03) for profiles
            pieces_needed = math.ceil((total_length * 1.03) / 12)
            profile_summary.append({
                'Profile': profile, 'Grade': grade, 
                'Total Length (m)': round(total_length, 2),
                '12m Pieces (3% waste)': pieces_needed
            })

    # Writing Result
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        pd.DataFrame(profile_summary).to_excel(writer, sheet_name='Profiles', index=False)
        pd.DataFrame(plate_summary).to_excel(writer, sheet_name='Plates', index=False)
    print(f"Successfully generated: {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", "-i", default="Tekla_List.xlsx")
    parser.add_argument("--output", "-o", default="Procurement_Plan.xlsx")
    args = parser.parse_args()
    process_procurement(args.input, args.output)
