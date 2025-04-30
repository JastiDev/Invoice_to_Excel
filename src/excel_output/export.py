"""
Excel output module for exporting data to Excel
"""
import pandas as pd


def export_to_excel(data, output_excel_path):
    """
    Export data to Excel with formatting
    
    Args:
        data (list): List of dictionaries containing invoice data
        output_excel_path (str): Path to save Excel file
    
    Returns:
        bool: True if successful, False otherwise
    """
    if not data:
        return False
        
    df = pd.DataFrame(data)
    
    # Calculate Tentative column
    df['Tentative'] = df.apply(
        lambda row: round((row['CostPerPacket'] / row['BarInParanthesis'] * 1.7), 2) 
        if row['BarInParanthesis'] > 0 
        else pd.NA, 
        axis=1
    )
    
    # Reorder columns to match the image format
    column_order = [
        'Purchased', 'Received', 'Code1', 'Code2', 'Brand', 'Description',
        'Product', 'CostPerPacket', 'TotalCost', 'BarInParanthesis', 'UnitCost', 'Tentative'
    ]
    
    # Make sure all columns exist before reordering
    for col in column_order:
        if col not in df.columns:
            df[col] = None
    
    df = df[column_order]
    
    # Apply Excel formatting
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Invoice Details')
        workbook = writer.book
        worksheet = writer.sheets['Invoice Details']
        
        # Format currency columns
        currency_columns = ['CostPerPacket', 'UnitCost', 'TotalCost', 'Tentative']
        for col in currency_columns:
            col_idx = column_order.index(col) + 1  # +1 because Excel columns start at 1
            for row in range(2, len(df) + 2):  # +2 to account for header and 1-based indexing
                cell = worksheet.cell(row=row, column=col_idx)
                cell.number_format = '#,##0.00'
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
    
    return True 