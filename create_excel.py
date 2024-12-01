import os

import pandas as pd
from openpyxl import load_workbook

def create_xlsx(pdf_data):
    # Clean up the first two pages (if needed)
    del pdf_data[0]
    del pdf_data[1]
    file_path = 'property_data.xlsx'
    # Convert the extracted data to a DataFrame
    df = pd.DataFrame.from_dict(pdf_data, orient='index')

    new_column_order = [
        'strata_plan',
        'lot',
        'unit_no',
        'levy_entitlement',
        'lot_street_name',
        'owner_name',
        'contact_number',
        'notice_address',
        'notice_address_street_address',
        'notice_address_suburb',
        'notice_address_state',
        'notice_address_postcode',
        'levy_address',
        'date_purchase',
        'date_entry',
        'owner_emails',
        'agent_name',
        'tenant_name',
        'tenant_contact',
        'vacant',
        'lease_start_date',
        'lease_end_date',
        'move_in_date',
        'review_date',
    ]
    # Reorder the DataFrame
    df = df[new_column_order]
    # Save the data to Excel
    df.to_excel(file_path, index=False)

    # Adjust column widths based on the content
    wb = load_workbook(file_path)
    ws = wb.active
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value is not None)
        adjusted_width = max_length + 2  # Add some padding for readability
        ws.column_dimensions[col[0].column_letter].width = adjusted_width

    # Save the modified Excel file
    wb.save('Data_from_pdf.xlsx')
    print(f"Data_from_pdf.xlsx has been created.")
    # Delete the file

    if os.path.exists(file_path):
        os.remove(file_path)
    else:
        pass