from invoice_entry import *

InvoiceEntry("etc")

ImportExportData.import_data()
InvoiceEntry.activate_window()

ImportExportData.import_header()
ImportExportData.import_vat()

ImportExportData.export_to_csv()
