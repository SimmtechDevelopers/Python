import invoice_entry

option = ["m365", "etc", "communication"]
invoice_entry.InvoiceEntry(option[0])

invoice_entry.ImportExportData.import_data()
invoice_entry.ImportExportData.export_to_csv()
