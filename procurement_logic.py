import pandas as pd 
import math 
import argparse 
import sys 
import re 
 
def process_procurement(input_file, output_file): 
    try: 
        if input_file.endswith('.csv'): 
            df = pd.read_csv(input_file) 
        else: 
            df = pd.read_excel(input_file) 
    except Exception as e: 
        print(f"Error reading {input_file}: {e}") 
        sys.exit(1) 
 
    required_columns = ['Profile', 'Grade', 'Length', 'Width', 'Quantity'] 
    for col in required_columns: 
        if col not in df.columns: 
            print(f"Error: Missing required column '{col}'") 
            sys.exit(1) 
 
    # Convert numeric columns 
    df['Length'] = pd.to_numeric(df['Length'], errors='coerce').fillna(0) 
    df['Width'] = pd.to_numeric(df['Width'], errors='coerce').fillna(0) 
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0) 
 
    # Unit detection logic
    # Check the average value of the 'Length' column.
    avg_length = df['Length'].mean()
    if avg_length > 100:
        # Assume units are in millimeters, normalize to meters
        print(f"Detected average length > 100 ({avg_length:.2f}). Normalizing Length and Width from mm to m.")
        df['Length'] = df['Length'] / 1000
        df['Width'] = df['Width'] / 1000
    else:
        print(f"Detected average length <= 100 ({avg_length:.2f}). Assuming units are already in meters.")

    # Separate Plates and Profiles 
    # Assuming Plates have Width > 0 and Profiles have Width == 0 or NaN 
    # A more robust way might be looking at Profile name (e.g., 'PL') 
    is_plate = df['Profile'].astype(str).str.upper().str.startswith('PL') | (df['Width'] > 0) 
     
    plates_df = df[is_plate].copy() 
    profiles_df = df[~is_plate].copy() 
 
    # --- Profile Logic --- 
    profile_summary = [] 
    if not profiles_df.empty: 
        # Group by Profile and Grade 
        profile_groups = profiles_df.groupby(['Profile', 'Grade']) 
        for (profile, grade), group in profile_groups: 
            # Length should be in meters, assuming input is in meters if values are small,  
            # but usually Tekla outputs in mm. Let's assume input is in meters as requested by division by 12. 
            # If the length is in mm (e.g., >1000 for standard parts), we should probably convert.  
            # The prompt doesn't specify unit conversion, so we'll assume Length is in meters or  
            # handle it transparently. Actually, let's just do Length * Quantity directly. 
            # The prompt says: "Calculate the total length needed. Divide by 12 meters (standard length), round up to the nearest integer, and add a 5% waste factor." 
            # Actually, standard formula with waste: math.ceil( (Total Length * 1.05) / 12 ) 
             
            total_length = (group['Length'] * group['Quantity']).sum() 
             
            # The exact phrasing "Calculate the total length needed. Divide by 12 meters (standard length), round up to the nearest integer, and add a 5% waste factor." 
            # Could mean: ceil(Total Length / 12) * 1.05 ? No, usually "add a 5% waste factor" means on the total length or on the pieces. 
            # "Divide by 12 meters (standard length), round up to the nearest integer, and add a 5% waste factor." 
            # Let's interpret as: pieces = math.ceil((total_length * 1.05) / 12) 
             
            pieces_needed = math.ceil((total_length * 1.05) / 12) 
             
            profile_summary.append({ 
                'Profile': profile, 
                'Grade': grade, 
                'Total Length Needed (m)': total_length, # assuming input is meters 
                '12m Standard Pieces (incl. 5% waste)': pieces_needed 
            }) 
 
    profile_summary_df = pd.DataFrame(profile_summary) 
 
 
    # --- Plate Logic --- 
    plate_summary = [] 
    if not plates_df.empty: 
        # We need Thickness. We can extract it from the 'Profile' column 
        # e.g., 'PL10' -> 10, 'PL 10*200' -> 10 
        def extract_thickness(profile_str): 
            match = re.search(r'\d+(\.\d+)?', str(profile_str)) 
            if match: 
                return float(match.group()) 
            return 0.0 
 
        plates_df['Thickness'] = plates_df['Profile'].apply(extract_thickness) 
         
        # Group by Thickness and Grade 
        plate_groups = plates_df.groupby(['Thickness', 'Grade']) 
        for (thickness, grade), group in plate_groups: 
            # Total area = Length * Width * Quantity 
            total_area = (group['Length'] * group['Width'] * group['Quantity']).sum() 
             
            # Divide by 18 sqm, add 10% waste, round up 
            # "Calculate how many standard plates are needed to cover the total area plus a 10% waste factor for nesting." 
            plates_needed = math.ceil((total_area * 1.10) / 18) 
             
            plate_summary.append({ 
                'Thickness': thickness, 
                'Grade': grade, 
                'Total Area (sqm)': total_area, 
                'Standard Plates (1.5x12m) (incl. 10% waste)': plates_needed 
            }) 
 
    plate_summary_df = pd.DataFrame(plate_summary) 
 
    # Write to Excel 
    try: 
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer: 
            if not profile_summary_df.empty: 
                profile_summary_df.to_excel(writer, sheet_name='Profiles', index=False) 
            if not plate_summary_df.empty: 
                plate_summary_df.to_excel(writer, sheet_name='Plates', index=False) 
            if profile_summary_df.empty and plate_summary_df.empty: 
                # Create a dummy sheet if both are empty 
                pd.DataFrame({'Message': ['No data processed']}).to_excel(writer, sheet_name='Summary', index=False) 
        print(f"Successfully wrote procurement plan to {output_file}") 
    except Exception as e: 
        print(f"Error writing to {output_file}: {e}") 
        sys.exit(1) 
 
if __name__ == "__main__": 
    parser = argparse.ArgumentParser(description="Process steel structure procurement from Tekla lists.") 
    parser.add_argument("--input", "-i", default="Tekla_List.xlsx", help="Input Excel/CSV file name") 
    parser.add_argument("--output", "-o", default="Procurement_Plan.xlsx", help="Output Excel file name") 
    args = parser.parse_args() 
     
    process_procurement(args.input, args.output)
