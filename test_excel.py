import openpyxl, os
from datetime import datetime

DATA_DIR = r"C:\laragon\www\TiffaniHume_Website\data"
CRM_FILE = os.path.join(DATA_DIR, "HUME_CRM.xlsx")

lead_data = {
    "name": "Test Write",
    "email": "test@example.com",
    "phone": "555-555-5555",
    "address": "123 Main Street",
    "notes": "Testing Excel append"
}

from app import log_lead_to_excel
log_lead_to_excel(lead_data)
