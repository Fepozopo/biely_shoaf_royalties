import openpyxl



def update_artist(artists, royalty_file, report_file):
    # Loop through artists
    for artist in artists:
        wb_royalty = openpyxl.load_workbook(royalty_file)
        ws_royalty = wb_royalty[f'{artist}']

        wb_report = openpyxl.load_workbook(report_file)
        ws_report = wb_report['Sheet1']

        # Define columns
        royalty_sku_col = 1 # 'A' column in the royalty file (1-based index)
        royalty_sales_col = 3 # 'C' column in the royalty file (1-based index)
        report_sku_col = 3 # 'C' column in the report (1-based index)
        report_sales_col = 6 # 'F' column in the report (1-based index)
        report_pointer = 1 # Pointer to the next row in the report

        print(f'Updating {artist}...')
        for row_royalty in range(3, ws_royalty.max_row + 1): # Loop through royalty rows
            royalty_sku = ws_royalty.cell(row=row_royalty, column=royalty_sku_col).value
            
            if royalty_sku is None:
                continue # Skip empty rows

            # Ensure royalty_sku is a string and strip leading/trailing whitespace
            royalty_sku = str(royalty_sku).strip()

            for row_report in range(report_pointer, ws_report.max_row + 1): # Loop through report rows
                report_sku = ws_report.cell(row=row_report, column=report_sku_col).value

                if report_sku is None:
                    continue # Skip empty rows

                # Ensure report_sku is a string and strip leading/trailing whitespace
                report_sku = str(report_sku).strip()

                if royalty_sku == report_sku:
                    report_sales_value = ws_report.cell(row=row_report, column=report_sales_col).value
                    
                    ws_royalty.cell(row=row_royalty, column=royalty_sales_col).value = report_sales_value # Update the royalty sales column with the report sales column
                    print(f"Updated SKU {royalty_sku} with sales value {report_sales_value}")
                    
                    report_pointer = row_report + 1 # Move the report pointer to the next row
                    break # Move to the next royalty row

        # Save and close the workbook
        wb_royalty.save(royalty_file)
        wb_report.close()
        wb_royalty.close()
