import pandas as pd
import os

def load_msku_mapping(file_path):
    """
    Load and process the MSKU to SKU mapping file.
    
    Args:
        file_path (str): Path to the MSKU to SKU mapping Excel file
        
    Returns:
        pd.DataFrame: Processed mapping dataframe
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Clean column names
        df.columns = df.columns.str.strip().str.lower()
        
        # Filter out rows with reference errors
        df = df[~df['status'].str.contains('reference error', case=False, na=False)]
        
        # Create a clean mapping dictionary
        msku_sku_map = {}
        for _, row in df.iterrows():
            msku = row['msku']
            sku = row['sku']
            status = row['status']
            
            if pd.notna(msku) and pd.notna(sku):
                msku_sku_map[sku] = {
                    'msku': msku,
                    'status': status
                }
        
        return msku_sku_map
    
    except Exception as e:
        print(f"Error loading MSKU mapping: {str(e)}")
        return None

def map_sku_to_msku(data_df, msku_map, sku_column='sku'):
    """
    Map SKUs to MSKUs in the data dataframe.
    
    Args:
        data_df (pd.DataFrame): DataFrame containing SKUs
        msku_map (dict): Dictionary mapping SKUs to MSKUs
        sku_column (str): Name of the SKU column in data_df
        
    Returns:
        pd.DataFrame: DataFrame with added MSKU information
    """
    try:
        # Create a copy of the dataframe
        df = data_df.copy()
        
        # Add MSKU and status columns
        df['msku'] = df[sku_column].map(lambda x: msku_map.get(x, {}).get('msku'))
        df['sku_status'] = df[sku_column].map(lambda x: msku_map.get(x, {}).get('status'))
        
        return df
    
    except Exception as e:
        print(f"Error mapping SKUs to MSKUs: {str(e)}")
        return data_df

if __name__ == "__main__":
    # Example usage
    msku_file = os.path.join("data", "msku_to_sku.xlsx")
    msku_map = load_msku_mapping(msku_file)
    
    if msku_map:
        print(f"Successfully loaded {len(msku_map)} SKU to MSKU mappings")
    else:
        print("Failed to load MSKU mapping") 