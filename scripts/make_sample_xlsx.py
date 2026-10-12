"""Generate public/sample_load_list.xlsx (run: python scripts/make_sample_xlsx.py)."""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

HEADERS = ["Description", "Type (DOL/VFD/Feeder)", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)"]
ROWS = [
    # description, type, amps, kW, voltage
    ("Boiler feed water pump", "DOL", None, 185, 433),
    ("Cooling water pump A", "DOL", None, 160, 433),
    ("Cooling water pump B (standby)", "DOL", None, 160, 433),
    ("ID fan", "VFD", None, 250, 433),
    ("FD fan", "VFD", None, 200, 433),
    ("Air compressor", "DOL", None, 132, 433),
    ("Conveyor 1", "DOL", None, 90, 433),
    ("Conveyor 2", "VFD", None, 75, 433),
    ("Crusher", "DOL", None, 315, 433),
    ("Dosing pump (amps only)", "DOL", 180, None, 433),
    ("Booster fan (amps only)", "VFD", 120, None, 433),
    ("Lighting & small power DB", "Feeder", 250, None, 433),
    ("MCC-2 incomer", "Feeder", 800, None, 433),
    ("Grinding mill", "DOL", None, 450, 690),
    ("Slurry pump", "VFD", None, 355, 690),
]

wb = Workbook()
ws = wb.active
ws.title = "Load list"
ws.append(HEADERS)
for r in ROWS:
    ws.append(list(r))
for c in ws[1]:
    c.font = Font(bold=True)
    c.alignment = Alignment(wrap_text=True, vertical="center")
for i, w in enumerate([34, 22, 16, 16, 16], start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"

notes = wb.create_sheet("Notes")
for line in [
    ("How the load list is read",),
    ("Only the first sheet ('Load list') is read. Keep the header row; column order does not matter.",),
    ("kW given -> motor of that kW. kVA = kW / (power factor x efficiency).",),
    ("Only amps given and Type is DOL or VFD -> motor. kVA = sqrt(3) x V x A.",),
    ("Only amps given and Type is anything else -> outgoing feeder. kVA = sqrt(3) x V x A.",),
    ("Power factor: DOL 0.7, VFD 0.9, feeder 0.7. Efficiency 95 %. All loads at 100 % (standby included).",),
    ("Voltage level: 433 or 690 (volts). Rows that do not match the selected transformer are listed as excluded.",),
    ("The values in this file are examples only.",),
]:
    notes.append(list(line))
notes["A1"].font = Font(bold=True)
notes.column_dimensions["A"].width = 110

wb.save("public/sample_load_list.xlsx")
print("saved")
