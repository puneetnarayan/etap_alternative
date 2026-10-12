"""Generate public/sample_load_list.xlsx (run: python scripts/make_sample_xlsx.py)."""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

HEADERS = ["Description", "Type (DOL/VFD/Feeder)", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)", "Working/Standby (W/S)", "Duty/standby group"]
ROWS = [
    # description, type, amps, kW, voltage, W/S, duty/standby group (optional)
    ("Boiler feed water pump A", "DOL", None, 185, 433, "W", "BFW"),
    ("Boiler feed water pump B", "DOL", None, 185, 433, "S", "BFW"),
    ("Cooling water pump A", "DOL", None, 160, 433, "W", "CWP"),
    ("Cooling water pump B", "DOL", None, 160, 433, "S", "CWP"),
    ("ID fan", "VFD", None, 250, 433, "W", None),
    ("FD fan", "VFD", None, 200, 433, "W", None),
    ("Air compressor", "DOL", None, 132, 433, "W", None),
    ("Conveyor 1", "DOL", None, 90, 433, "W", None),
    ("Conveyor 2", "VFD", None, 75, 433, "W", None),
    ("Crusher", "DOL", None, 315, 433, "W", None),
    ("Dosing pump (amps only)", "DOL", 180, None, 433, "W", None),
    ("Booster fan (amps only)", "VFD", 120, None, 433, "W", None),
    ("Lighting & small power DB", "Feeder", 250, None, 433, "W", None),
    ("MCC-2 incomer", "Feeder", 800, None, 433, "W", None),
    ("Grinding mill", "DOL", None, 450, 690, "W", None),
    ("Slurry pump", "VFD", None, 355, 690, "W", None),
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
for i, w in enumerate([30, 22, 16, 16, 16, 22, 20], start=1):
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
    ("Working/Standby (optional): W = working, S = standby. Shown as W / S in the results. Standby loads are counted at 100 % like working loads.",),
    ("Duty/standby group (optional): loads with the same group name are placed on different transformers' bus sections.",),
    ("Optional columns can be left out; extra columns are ignored.",),
    ("The values in this file are examples only.",),
]:
    notes.append(list(line))
notes["A1"].font = Font(bold=True)
notes.column_dimensions["A"].width = 110

wb.save("public/sample_load_list.xlsx")
print("saved")
