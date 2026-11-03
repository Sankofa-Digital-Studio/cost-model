"""
UNIVERSAL DEMOLITION, REHABILITATION & CIVIL COSTING WORKBOOK GENERATOR
=======================================================================
Engineered for Senatla Trading Enterprise / Sankofa Digital Studio.
Produces 100% dynamic formula-driven executive workbooks matching
West Gold Plant & CAPM Tau Mine #10 executive standards.
"""

import os
import json
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# DEFAULT PROJECT PRESETS
# =============================================================================

PRESET_WEST_GOLD = {
    "project_name": "West Gold Plant Demolition & Rehabilitation Model",
    "client_name": "West Gold Plant Operations / Senatla Trading",
    "contract_ref": "WGP-DEMO-REHAB-2026-01",
    "duration_months": 12,
    "output_filename": "West_Gold_Plant_Universal_Model.xlsx",
    "rates": {
        "site_est": 35000.0,
        "steel_dismantle": 21000.0,
        "concrete_breakup": 240.0,
        "rubble_haul": 120.0,
        "shaping_levelling": 18.50,
        "concrete_density": 2.40,
        "pg_pct": 0.10,
        "contingency_pct": 0.10,
        "profit_pct": 0.05,
        "vat_pct": 0.15
    },
    "phases": [
        {"name": "Phase 1 - Smelt House", "footprint_m2": 32000, "structures": 4, "concrete_m3": 1800, "duration_months": 2},
        {"name": "Phase 2 - Thickener Dams", "footprint_m2": 45000, "structures": 2, "concrete_m3": 3500, "duration_months": 2},
        {"name": "Phase 3 - Mill Section", "footprint_m2": 28000, "structures": 6, "concrete_m3": 4200, "duration_months": 3},
        {"name": "Phase 4 - Silos & Fdns", "footprint_m2": 15000, "structures": 3, "concrete_m3": 2900, "duration_months": 2},
        {"name": "Phase 5 - Workshops", "footprint_m2": 22000, "structures": 5, "concrete_m3": 1400, "duration_months": 2},
        {"name": "Phase 6 - Conveyor Struct", "footprint_m2": 18000, "structures": 4, "concrete_m3": 950, "duration_months": 2},
        {"name": "Phase 7 - Main Gate", "footprint_m2": 12000, "structures": 2, "concrete_m3": 600, "duration_months": 1}
    ]
}

PRESET_SENATLA_VILLAGES = {
    "project_name": "Senatla 5-Village Demolition & Rubble Haulage Model",
    "client_name": "Senatla Trading Enterprise 26 (Pty) Ltd",
    "contract_ref": "STE26-VILLAGE-DEMO-REHAB-2026",
    "duration_months": 6,
    "output_filename": "Senatla_5_Village_Demolition_Rehabilitation_Model.xlsx",
    "rates": {
        "site_est": 35000.0,
        "steel_dismantle": 21000.0,
        "concrete_breakup": 240.0,
        "rubble_haul": 120.0,
        "shaping_levelling": 18.50,
        "concrete_density": 2.40,
        "pg_pct": 0.10,
        "contingency_pct": 0.10,
        "profit_pct": 0.05,
        "vat_pct": 0.15
    },
    "phases": [
        {"name": "Boiketlong Village", "footprint_m2": 15000, "structures": 2, "concrete_m3": 384, "rubble_tons": 922, "struct_rate": 21000.0, "duration_months": 1},
        {"name": "Polokong Village", "footprint_m2": 25000, "structures": 3, "concrete_m3": 962, "rubble_tons": 2310, "struct_rate": 21000.0, "duration_months": 2},
        {"name": "Woolaways Village", "footprint_m2": 14000, "structures": 2, "concrete_m3": 385, "rubble_tons": 924, "struct_rate": 21000.0, "duration_months": 1},
        {"name": "Jabulani Village", "footprint_m2": 26000, "structures": 5, "concrete_m3": 962, "rubble_tons": 2310, "struct_rate": 21000.0, "duration_months": 2},
        {"name": "Pabalong Village", "footprint_m2": 45000, "structures": 8, "concrete_m3": 2496, "rubble_tons": 5992, "struct_rate": 19500.0, "duration_months": 3}
    ]
}


def build_universal_costing_workbook(project_config: dict, output_path: str = None) -> str:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Colors
    NAVY_DARK = "1B365D"       # Primary Headers (#1B365D)
    NAVY_MED = "2E5B88"        # Subheaders (#2E5B88)
    GOLD_ACCENT = "C5A059"     # Accent (#C5A059)
    GRAY_LIGHT = "F4F6F9"      # Zebra striping
    GRAY_SUBTOTAL = "E9EEF4"   # Subtotal row fill
    GRAY_BORDER = "D9D9D9"     # Borders
    INPUT_YELLOW = "FFF2CC"    # Input cells
    INPUT_BORDER = "D4AF37"    # Input borders
    PROFIT_GREEN_BORDER = "274E13"
    PROFIT_RED_TEXT = "C00000"
    WHITE = "FFFFFF"

    # Fonts
    font_main_title = Font(name="Segoe UI", size=14, bold=True, color=WHITE)
    font_sub_title = Font(name="Segoe UI", size=9, italic=True, color=WHITE)
    font_sec_hdr = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    font_tbl_hdr = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
    font_data = Font(name="Segoe UI", size=10, color="000000")
    font_data_bold = Font(name="Segoe UI", size=10, bold=True, color="000000")
    font_subtotal = Font(name="Segoe UI", size=10, bold=True, color="1B365D")
    font_grandtotal = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    font_input = Font(name="Segoe UI", size=10, bold=True, color="002060")
    font_profit_pct = Font(name="Segoe UI", size=12, bold=True, color=PROFIT_RED_TEXT)
    font_kpi_num = Font(name="Segoe UI", size=13, bold=True, color="1B365D")
    font_kpi_lbl = Font(name="Segoe UI", size=8, bold=True, color="595959")
    font_note = Font(name="Segoe UI", size=9, italic=True, color="595959")

    # Fills
    fill_main_title = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
    fill_sec_hdr = PatternFill(start_color=NAVY_MED, end_color=NAVY_MED, fill_type="solid")
    fill_tbl_hdr = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
    fill_zebra = PatternFill(start_color=GRAY_LIGHT, end_color=GRAY_LIGHT, fill_type="solid")
    fill_subtotal = PatternFill(start_color=GRAY_SUBTOTAL, end_color=GRAY_SUBTOTAL, fill_type="solid")
    fill_grandtotal = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
    fill_input = PatternFill(start_color=INPUT_YELLOW, end_color=INPUT_YELLOW, fill_type="solid")
    fill_kpi = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")

    # Borders
    thin_side = Side(style='thin', color=GRAY_BORDER)
    med_navy_side = Side(style='medium', color=NAVY_DARK)
    double_navy_side = Side(style='double', color=NAVY_DARK)
    gold_side = Side(style='thin', color=INPUT_BORDER)
    profit_border_side = Side(style='medium', color=PROFIT_GREEN_BORDER)

    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_subtotal = Border(left=thin_side, right=thin_side, top=Side(style='thin', color=NAVY_DARK), bottom=double_navy_side)
    border_grandtotal = Border(left=med_navy_side, right=med_navy_side, top=Side(style='thin', color=WHITE), bottom=Side(style='double', color=WHITE))
    border_input = Border(left=gold_side, right=gold_side, top=gold_side, bottom=gold_side)
    border_profit_box = Border(left=profit_border_side, right=profit_border_side, top=profit_border_side, bottom=profit_border_side)
    border_kpi = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Number Formats
    FMT_CURRENCY = 'R #,##0.00'
    FMT_QTY = '#,##0'
    FMT_QTY_DEC = '#,##0.00'
    FMT_PERCENT = '0.0%'

    proj_name = project_config.get("project_name", "Demolition & Rehabilitation Financial Model")
    client_name = project_config.get("client_name", "Senatla Trading Enterprise 26 (Pty) Ltd")
    contract_ref = project_config.get("contract_ref", "CONTRACT-2026")
    duration_months = int(project_config.get("duration_months", 12))
    rates_cfg = project_config.get("rates", {})
    phases = project_config.get("phases", [])

    if not output_path:
        output_path = project_config.get("output_filename", "Universal_Costing_Model.xlsx")

    # =========================================================================
    # SHEET 1: RATES & PARAMETERS (MASTER CONTROL REPOSITORY)
    # =========================================================================
    ws_rates = wb.create_sheet(title="Rates & Parameters")
    ws_rates.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_rates.merge_cells("A1:G1")
    ws_rates["A1"] = f"{proj_name.upper()} — MASTER RATES & COMMERCIAL PARAMETERS"
    ws_rates["A1"].font = font_main_title
    ws_rates["A1"].fill = fill_main_title
    ws_rates["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_rates.row_dimensions[1].height = 36

    ws_rates.merge_cells("A2:G2")
    ws_rates["A2"] = f"Client: {client_name} | Contract Ref: {contract_ref} | Program Duration: {duration_months} Months"
    ws_rates["A2"].font = font_sub_title
    ws_rates["A2"].fill = fill_main_title
    ws_rates["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_rates.row_dimensions[2].height = 18

    # Section 1: Standard Unit Rates
    ws_rates.merge_cells("A4:G4")
    ws_rates["A4"] = "1. STANDARD OPERATIONAL UNIT RATES (BASELINE DIRECT RATES)"
    ws_rates["A4"].font = font_sec_hdr
    ws_rates["A4"].fill = fill_sec_hdr
    ws_rates["A4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_rates.row_dimensions[4].height = 24

    rate_hdrs = ["Rate Code", "Activity Description", "Unit of Measure (UOM)", "Standard Unit Rate (ZAR)", "Rate Category", "Notes / Specification", "Editable Tag"]
    for c_idx, h in enumerate(rate_hdrs, 1):
        cell = ws_rates.cell(5, c_idx, h)
        cell.font = font_tbl_hdr
        cell.fill = fill_tbl_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_cell
    ws_rates.row_dimensions[5].height = 26

    std_rates = [
        ("RATE_EST", "Site Establishment, mobilization & safety compliance", "Lot", rates_cfg.get("site_est", 35000.0), "Establishment", "Fixed cost per site / zone work package", "INPUT"),
        ("RATE_STEEL_STD", "Cut up and remove structural steel frames / trusses", "R/struct", rates_cfg.get("steel_dismantle", 21000.0), "Dismantling", "Standard structural steel framework per structure", "INPUT"),
        ("RATE_STEEL_HVY", "Heavy structural steel / headgear / gantry dismantling", "R/struct", rates_cfg.get("steel_heavy", 28500.0), "Specialist", "Industrial heavy steel and high-level frames", "INPUT"),
        ("RATE_STEEL_RES", "Residential / hostel / masonry steel framing removal", "R/struct", rates_cfg.get("steel_residential", 19500.0), "Residential", "Light steel roofing, trusses and cladding", "INPUT"),
        ("RATE_CONC_STD", "Breakup concrete slabs and foundations (±1.5m below NGL)", "m³", rates_cfg.get("concrete_breakup", 240.0), "Demolition", "Hydraulic hammer breakup and pulverization", "INPUT"),
        ("RATE_CONC_SILO", "Breakup heavy concrete silos and plinths", "m³", rates_cfg.get("concrete_silo", 260.0), "Demolition", "High-strength structural concrete demolition", "INPUT"),
        ("RATE_HAUL_RUBBLE", "Loading and hauling clean rubble to designated dump/stockpile", "R/ton", rates_cfg.get("rubble_haul", 120.0), "Haulage", "Transporting clean pulverized building rubble", "INPUT"),
        ("RATE_SHAPE_LEVEL", "Shaping, levelling and cleaning off area footprint", "m²", rates_cfg.get("shaping_levelling", 18.50), "Rehabilitation", "Final earthworks shaping to natural contour", "INPUT"),
        ("RATE_WRD_CRUSH", "Waste rock dump processing: mobile crushing & sizing", "ton", rates_cfg.get("wrd_crush", 45.0), "Processing", "Sizing fill for backfill or roadbase", "INPUT"),
        ("RATE_SHAFT_FILL", "Transport and stockpile material for deep shaft filling", "ton", rates_cfg.get("shaft_fill", 55.0), "Shaft Package", "High-capacity continuous shaft filling", "INPUT"),
        ("RATE_CAP_CONC", "Reinforced concrete collar capping slab (40 MPa certified)", "m³", rates_cfg.get("capping_concrete", 4200.0), "Structural Civil", "Cast-in-place high strength capping slab", "INPUT"),
    ]

    r_start = 6
    for idx, (code, desc, uom, val, cat, notes, tag) in enumerate(std_rates):
        r = r_start + idx
        ws_rates.cell(r, 1, code).alignment = Alignment(horizontal="center")
        ws_rates.cell(r, 2, desc).alignment = Alignment(horizontal="left")
        ws_rates.cell(r, 3, uom).alignment = Alignment(horizontal="center")
        
        c_val = ws_rates.cell(r, 4, val)
        c_val.number_format = FMT_CURRENCY
        c_val.alignment = Alignment(horizontal="right")
        c_val.font = font_input
        c_val.fill = fill_input
        c_val.border = border_input

        ws_rates.cell(r, 5, cat).alignment = Alignment(horizontal="center")
        ws_rates.cell(r, 6, notes).alignment = Alignment(horizontal="left")
        ws_rates.cell(r, 7, tag).alignment = Alignment(horizontal="center")

        for col_i in [1, 2, 3, 5, 6, 7]:
            c = ws_rates.cell(r, col_i)
            c.font = font_data
            c.border = border_cell
            if idx % 2 == 1:
                c.fill = fill_zebra
        ws_rates.row_dimensions[r].height = 20

    # Section 2: Master Commercial Multipliers
    r_multi_hdr = r_start + len(std_rates) + 2
    ws_rates.merge_cells(f"A{r_multi_hdr}:G{r_multi_hdr}")
    ws_rates[f"A{r_multi_hdr}"] = "2. MASTER COMMERCIAL & CONTRACT FINANCIAL MULTIPLIERS"
    ws_rates[f"A{r_multi_hdr}"].font = font_sec_hdr
    ws_rates[f"A{r_multi_hdr}"].fill = fill_sec_hdr
    ws_rates[f"A{r_multi_hdr}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_rates.row_dimensions[r_multi_hdr].height = 24

    for c_idx, h in enumerate(["Parameter Code", "Commercial Parameter Description", "Unit", "Value / Percentage", "Category", "Contract Reference", "Editable Tag"], 1):
        cell = ws_rates.cell(r_multi_hdr + 1, c_idx, h)
        cell.font = font_tbl_hdr
        cell.fill = fill_tbl_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_cell
    ws_rates.row_dimensions[r_multi_hdr + 1].height = 24

    commercial_params = [
        ("PARAM_PG_PCT", "Preliminaries and General (P&G) Allowance", "%", rates_cfg.get("pg_pct", 0.10), "Overhead", "Contract Preliminaries & General allowance", "INPUT", FMT_PERCENT),
        ("PARAM_CONT_PCT", "Contract Contingencies Allowance", "%", rates_cfg.get("contingency_pct", 0.10), "Risk", "Unforeseen physical site conditions contingency", "INPUT", FMT_PERCENT),
        ("PARAM_PROFIT_PCT", "Contractor Net Operating Profit Margin", "%", rates_cfg.get("profit_pct", 0.05), "Commercial", "Net return on project execution spend", "INPUT", FMT_PERCENT),
        ("PARAM_VAT_PCT", "Statutory SARS Value Added Tax (VAT)", "%", rates_cfg.get("vat_pct", 0.15), "Tax", "South African statutory standard VAT rate", "INPUT", FMT_PERCENT),
        ("PARAM_DENSITY_CONC", "Concrete Rubble Bulking Density Factor", "t/m³", rates_cfg.get("concrete_density", 2.40), "Geotechnical", "Tonnage conversion factor per m³ of solid concrete", "INPUT", FMT_QTY_DEC),
        ("PARAM_DENSITY_FILL", "Bulk Soil / Waste Rock Fill Density Factor", "t/m³", rates_cfg.get("fill_density", 1.70), "Geotechnical", "In-situ bulk density conversion per m³ of fill", "INPUT", FMT_QTY_DEC),
        ("PARAM_CLEAR_DEPTH", "Standard Demolition Depth Below NGL", "m", 1.50, "Engineering", "Contractual clearance specification below surface", "INPUT", FMT_QTY_DEC),
    ]

    r_p_start = r_multi_hdr + 2
    for idx, (code, desc, uom, val, cat, ref, tag, fmt) in enumerate(commercial_params):
        r = r_p_start + idx
        ws_rates.cell(r, 1, code).alignment = Alignment(horizontal="center")
        ws_rates.cell(r, 2, desc).alignment = Alignment(horizontal="left")
        ws_rates.cell(r, 3, uom).alignment = Alignment(horizontal="center")

        c_val = ws_rates.cell(r, 4, val)
        c_val.number_format = fmt
        c_val.alignment = Alignment(horizontal="right")
        c_val.font = font_input
        c_val.fill = fill_input
        c_val.border = border_input

        ws_rates.cell(r, 5, cat).alignment = Alignment(horizontal="center")
        ws_rates.cell(r, 6, ref).alignment = Alignment(horizontal="left")
        ws_rates.cell(r, 7, tag).alignment = Alignment(horizontal="center")

        for col_i in [1, 2, 3, 5, 6, 7]:
            c = ws_rates.cell(r, col_i)
            c.font = font_data
            c.border = border_cell
            if idx % 2 == 1:
                c.fill = fill_zebra
        ws_rates.row_dimensions[r].height = 20

    # Auto-fit column widths
    ws_rates.column_dimensions["A"].width = 22
    ws_rates.column_dimensions["B"].width = 46
    ws_rates.column_dimensions["C"].width = 16
    ws_rates.column_dimensions["D"].width = 24
    ws_rates.column_dimensions["E"].width = 20
    ws_rates.column_dimensions["F"].width = 38
    ws_rates.column_dimensions["G"].width = 16

    # =========================================================================
    # SHEET 2: MASTER BILL OF QUANTITIES (BOQ BREAKDOWN)
    # =========================================================================
    ws_boq = wb.create_sheet(title="Master BOQ Breakdown")
    ws_boq.views.sheetView[0].showGridLines = True

    ws_boq.merge_cells("A1:G1")
    ws_boq["A1"] = f"{proj_name.upper()} — MASTER BILL OF QUANTITIES (BOQ)"
    ws_boq["A1"].font = font_main_title
    ws_boq["A1"].fill = fill_main_title
    ws_boq["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_boq.row_dimensions[1].height = 36

    ws_boq.merge_cells("A2:G2")
    ws_boq["A2"] = "Consolidated Work Package Quantities, Parameter-Linked Unit Rates & Direct Cost Extensions"
    ws_boq["A2"].font = font_sub_title
    ws_boq["A2"].fill = fill_main_title
    ws_boq["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_boq.row_dimensions[2].height = 18

    boq_hdrs = ["Item Ref", "Work Package / Activity Description", "Unit", "Quantity", "Unit Rate (ZAR)", "Total Direct Amount (ZAR)", "Rate Link / Source"]
    for c_idx, h in enumerate(boq_hdrs, 1):
        cell = ws_boq.cell(4, c_idx, h)
        cell.font = font_tbl_hdr
        cell.fill = fill_tbl_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_cell
    ws_boq.row_dimensions[4].height = 26

    current_boq_row = 5
    phase_subtotal_cells = []

    for p_idx, phase in enumerate(phases, 1):
        p_name = phase.get("name", f"Phase {p_idx}")
        area_m2 = phase.get("footprint_m2", 10000)
        structs = phase.get("structures", 2)
        conc_m3 = phase.get("concrete_m3", 500)
        rubble_tons = phase.get("rubble_tons", round(conc_m3 * rates_cfg.get("concrete_density", 2.40)))
        struct_rate_val = phase.get("struct_rate", rates_cfg.get("steel_dismantle", 21000.0))

        # Phase Header Row
        ws_boq.merge_cells(f"A{current_boq_row}:G{current_boq_row}")
        ws_boq[f"A{current_boq_row}"] = f"{p_idx}.0  {p_name.upper()} (FOOTPRINT: {area_m2:,.0f} m²)"
        ws_boq[f"A{current_boq_row}"].font = font_sec_hdr
        ws_boq[f"A{current_boq_row}"].fill = fill_sec_hdr
        ws_boq[f"A{current_boq_row}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws_boq.row_dimensions[current_boq_row].height = 22
        current_boq_row += 1

        p_start_row = current_boq_row

        items = [
            (f"{p_idx}.1", f"Site Establishment, mobilization & safety perimeter ({p_name})", "Lot", 1, "='Rates & Parameters'!$D$6", "'Rates & Parameters'!D6"),
            (f"{p_idx}.2", f"Dismantling of structural steel frames, roofing & trusses ({p_name})", "R/struct", structs, f"='Rates & Parameters'!$D$7" if struct_rate_val == rates_cfg.get("steel_dismantle", 21000.0) else struct_rate_val, "'Rates & Parameters'!D7"),
            (f"{p_idx}.3", f"Breakup concrete foundations and slabs ±1.5m below NGL ({p_name})", "m³", conc_m3, "='Rates & Parameters'!$D$10", "'Rates & Parameters'!D10"),
            (f"{p_idx}.4", f"Remove / Load / Haul clean rubble to designated stockpile ({p_name})", "R/ton", rubble_tons, "='Rates & Parameters'!$D$12", "'Rates & Parameters'!D12"),
            (f"{p_idx}.5", f"Shaping, levelling and cleaning off area ({p_name})", "m²", area_m2, "='Rates & Parameters'!$D$13", "'Rates & Parameters'!D13"),
        ]

        for item_ref, desc, uom, qty, rate_expr, src_note in items:
            ws_boq.cell(current_boq_row, 1, item_ref).alignment = Alignment(horizontal="center")
            ws_boq.cell(current_boq_row, 2, desc).alignment = Alignment(horizontal="left")
            ws_boq.cell(current_boq_row, 3, uom).alignment = Alignment(horizontal="center")
            
            c_qty = ws_boq.cell(current_boq_row, 4, qty)
            c_qty.number_format = FMT_QTY
            c_qty.alignment = Alignment(horizontal="right")

            c_rate = ws_boq.cell(current_boq_row, 5, rate_expr)
            c_rate.number_format = FMT_CURRENCY
            c_rate.alignment = Alignment(horizontal="right")

            c_total = ws_boq.cell(current_boq_row, 6, f"=D{current_boq_row}*E{current_boq_row}")
            c_total.number_format = FMT_CURRENCY
            c_total.alignment = Alignment(horizontal="right")
            c_total.font = font_data_bold

            ws_boq.cell(current_boq_row, 7, src_note).alignment = Alignment(horizontal="center")

            for col_i in range(1, 8):
                c = ws_boq.cell(current_boq_row, col_i)
                c.border = border_cell
                if col_i not in [4, 5, 6]:
                    c.font = font_data
            ws_boq.row_dimensions[current_boq_row].height = 20
            current_boq_row += 1

        # Phase Subtotal Row
        p_end_row = current_boq_row - 1
        ws_boq.merge_cells(f"A{current_boq_row}:E{current_boq_row}")
        ws_boq[f"A{current_boq_row}"] = f"SUB-TOTAL {p_name.upper()} (DIRECT EXPENDITURE)"
        ws_boq[f"A{current_boq_row}"].font = font_subtotal
        ws_boq[f"A{current_boq_row}"].alignment = Alignment(horizontal="right", vertical="center")
        
        c_sub = ws_boq.cell(current_boq_row, 6, f"=SUM(F{p_start_row}:F{p_end_row})")
        c_sub.number_format = FMT_CURRENCY
        c_sub.alignment = Alignment(horizontal="right")
        c_sub.font = font_subtotal

        phase_subtotal_cells.append(f"F{current_boq_row}")

        for col_i in range(1, 8):
            cell = ws_boq.cell(current_boq_row, col_i)
            cell.fill = fill_subtotal
            cell.border = border_subtotal
        ws_boq.row_dimensions[current_boq_row].height = 22
        current_boq_row += 2

    # Grand Total Row
    ws_boq.merge_cells(f"A{current_boq_row}:E{current_boq_row}")
    ws_boq[f"A{current_boq_row}"] = "TOTAL DIRECT PROJECT EXPENDITURE (EXCL. P&G, CONTINGENCIES & VAT)"
    ws_boq[f"A{current_boq_row}"].font = font_grandtotal
    ws_boq[f"A{current_boq_row}"].alignment = Alignment(horizontal="right", vertical="center")

    sum_expr = "+".join(phase_subtotal_cells)
    c_grand = ws_boq.cell(current_boq_row, 6, f"={sum_expr}")
    c_grand.number_format = FMT_CURRENCY
    c_grand.alignment = Alignment(horizontal="right")
    c_grand.font = font_grandtotal

    boq_grand_total_coord = f"'Master BOQ Breakdown'!F{current_boq_row}"

    for col_i in range(1, 8):
        cell = ws_boq.cell(current_boq_row, col_i)
        cell.fill = fill_grandtotal
        cell.border = border_grandtotal
    ws_boq.row_dimensions[current_boq_row].height = 26

    ws_boq.column_dimensions["A"].width = 14
    ws_boq.column_dimensions["B"].width = 54
    ws_boq.column_dimensions["C"].width = 14
    ws_boq.column_dimensions["D"].width = 16
    ws_boq.column_dimensions["E"].width = 22
    ws_boq.column_dimensions["F"].width = 26
    ws_boq.column_dimensions["G"].width = 24

    # =========================================================================
    # SHEET 3: EXECUTIVE SUMMARY & FINANCIALS (WATERFALL & PROFIT MODEL)
    # =========================================================================
    ws_sum = wb.create_sheet(title="Executive Summary & Financials")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.merge_cells("A1:G1")
    ws_sum["A1"] = f"{proj_name.upper()} — EXECUTIVE FINANCIAL SUMMARY & TENDER WATERFALL"
    ws_sum["A1"].font = font_main_title
    ws_sum["A1"].fill = fill_main_title
    ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 36

    ws_sum.merge_cells("A2:G2")
    ws_sum["A2"] = f"Client: {client_name} | Total Execution Period: {duration_months} Calendar Months"
    ws_sum["A2"].font = font_sub_title
    ws_sum["A2"].fill = fill_main_title
    ws_sum["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 18

    # KPI Top Cards (Row 4 & 5)
    kpis = [
        ("TOTAL DIRECT SPEND", f"={boq_grand_total_coord}", FMT_CURRENCY, "Consolidated Phase BOQ"),
        ("TENDER PRICE (EXCL VAT)", "=E21", FMT_CURRENCY, "Project Cost A + Profit"),
        ("TOTAL CONTRACT PRICE", "=E23", FMT_CURRENCY, "Subtotal B + 15% VAT"),
        ("MONTHLY RUN-RATE", f"=E17/{duration_months}", FMT_CURRENCY, f"{duration_months} Month Execution Period"),
    ]

    card_cols = [("A", "B"), ("C", "D"), ("E", "E"), ("F", "G")]
    for idx, (lbl, formula_val, fmt, subtext) in enumerate(kpis):
        c_left, c_right = card_cols[idx]
        if c_left != c_right:
            ws_sum.merge_cells(f"{c_left}4:{c_right}4")
            ws_sum.merge_cells(f"{c_left}5:{c_right}5")

        top_cell = ws_sum[f"{c_left}4"]
        top_cell.value = lbl
        top_cell.font = font_kpi_lbl
        top_cell.alignment = Alignment(horizontal="center", vertical="center")
        top_cell.fill = fill_kpi

        val_cell = ws_sum[f"{c_left}5"]
        val_cell.value = formula_val
        val_cell.number_format = fmt
        val_cell.font = font_kpi_num
        val_cell.alignment = Alignment(horizontal="center", vertical="center")
        val_cell.fill = fill_kpi

        for r_k in [4, 5]:
            col_start_idx = openpyxl.utils.column_index_from_string(c_left)
            col_end_idx = openpyxl.utils.column_index_from_string(c_right)
            for c_i in range(col_start_idx, col_end_idx + 1):
                ws_sum.cell(r_k, c_i).border = border_kpi

    ws_sum.row_dimensions[4].height = 18
    ws_sum.row_dimensions[5].height = 26

    # Commercial Waterfall Table
    ws_sum.merge_cells("A7:G7")
    ws_sum["A7"] = "COMMERCIAL COST WATERFALL & CONTRACT PRICING MODEL"
    ws_sum["A7"].font = font_sec_hdr
    ws_sum["A7"].fill = fill_sec_hdr
    ws_sum["A7"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_sum.row_dimensions[7].height = 24

    for c_idx, h in enumerate(["Line Ref", "Cost Element / Commercial Level", "Formula / Basis Description", "Rate / Multiplier Link", "Commercial Amount (ZAR)", "Contract %", "Review Notes"], 1):
        cell = ws_sum.cell(8, c_idx, h)
        cell.font = font_tbl_hdr
        cell.fill = fill_tbl_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_cell
    ws_sum.row_dimensions[8].height = 24

    # Waterfall Rows:
    # Row 9: Direct Spend (from BOQ Grand Total)
    # Row 10: Preliminaries & General (10%)
    # Row 11: Subtotal (Excl. Contingencies)
    # Row 12: Contingencies (10%)
    # Row 13: Empty separator
    # Row 14: Subtotal: Inclusive of P&G and Contingencies
    # Row 15: Blank
    # Row 16: Profit Percentage (5%)
    # Row 17: Project Costs A (Excl VAT)
    # Row 18: Monthly Run-Rate
    # Row 19: Net Profit (5%)
    # Row 20: Blank
    # Row 21: Sub-Total B Tender Price (Excl. VAT)
    # Row 22: SARS VAT (15%)
    # Row 23: TOTAL TENDER PRICE C (INCL. VAT)

    ws_sum.cell(9, 1, "DIR_01").alignment = Alignment(horizontal="center")
    ws_sum.cell(9, 2, "Consolidated Direct Operational Expenditure").alignment = Alignment(horizontal="left")
    ws_sum.cell(9, 3, "Sum of all phase work packages and civil demolition").alignment = Alignment(horizontal="left")
    ws_sum.cell(9, 4, "='Master BOQ Breakdown' Schedule").alignment = Alignment(horizontal="center")
    c9 = ws_sum.cell(9, 5, f"={boq_grand_total_coord}")
    c9.number_format = FMT_CURRENCY
    c9.font = font_data_bold
    c9.alignment = Alignment(horizontal="right")
    ws_sum.cell(9, 6, "=E9/$E$21").number_format = FMT_PERCENT
    ws_sum.cell(9, 7, "Direct baseline costs").alignment = Alignment(horizontal="center")

    ws_sum.cell(10, 1, "PG_10").alignment = Alignment(horizontal="center")
    ws_sum.cell(10, 2, "Preliminaries and General (10%)").alignment = Alignment(horizontal="left")
    ws_sum.cell(10, 3, "= Direct Subtotal * P&G Allowance Rate").alignment = Alignment(horizontal="left")
    ws_sum.cell(10, 4, "='Rates & Parameters'!$D$19").alignment = Alignment(horizontal="center")
    c10 = ws_sum.cell(10, 5, "=E9*'Rates & Parameters'!$D$19")
    c10.number_format = FMT_CURRENCY
    c10.font = font_data_bold
    c10.alignment = Alignment(horizontal="right")
    ws_sum.cell(10, 6, "=E10/$E$21").number_format = FMT_PERCENT
    ws_sum.cell(10, 7, "Master Parameter Linked").alignment = Alignment(horizontal="center")

    ws_sum.cell(11, 1, "SUB_01").alignment = Alignment(horizontal="center")
    ws_sum.cell(11, 2, "Sub-Total (Excluding Contingencies)").alignment = Alignment(horizontal="left")
    ws_sum.cell(11, 3, "= Direct Operational Spend + Preliminaries & General").alignment = Alignment(horizontal="left")
    ws_sum.cell(11, 4, "= E9 + E10").alignment = Alignment(horizontal="center")
    c11 = ws_sum.cell(11, 5, "=E9+E10")
    c11.number_format = FMT_CURRENCY
    c11.font = font_subtotal
    c11.alignment = Alignment(horizontal="right")
    ws_sum.cell(11, 6, "=E11/$E$21").number_format = FMT_PERCENT
    ws_sum.cell(11, 7, "Subtotal Level 1").alignment = Alignment(horizontal="center")

    ws_sum.cell(12, 1, "CONT_10").alignment = Alignment(horizontal="center")
    ws_sum.cell(12, 2, "Contingencies Allowance (10%)").alignment = Alignment(horizontal="left")
    ws_sum.cell(12, 3, "= Subtotal Excl. Contingencies * Contingency Rate").alignment = Alignment(horizontal="left")
    ws_sum.cell(12, 4, "='Rates & Parameters'!$D$20").alignment = Alignment(horizontal="center")
    c12 = ws_sum.cell(12, 5, "=E11*'Rates & Parameters'!$D$20")
    c12.number_format = FMT_CURRENCY
    c12.font = font_data_bold
    c12.alignment = Alignment(horizontal="right")
    ws_sum.cell(12, 6, "=E12/$E$21").number_format = FMT_PERCENT
    ws_sum.cell(12, 7, "Master Parameter Linked").alignment = Alignment(horizontal="center")

    # Sub-Total Level 2 (Inclusive of P&G and Contingencies)
    ws_sum.cell(14, 1, "SUB_02").alignment = Alignment(horizontal="center")
    ws_sum.cell(14, 2, "Sub-Total: Inclusive of Preliminaries and General & Contingencies").alignment = Alignment(horizontal="left")
    ws_sum.cell(14, 3, "= Subtotal Excl. Contingencies + Contingencies Allowance").alignment = Alignment(horizontal="left")
    ws_sum.cell(14, 4, "= E11 + E12").alignment = Alignment(horizontal="center")
    c14 = ws_sum.cell(14, 5, "=E11+E12")
    c14.number_format = FMT_CURRENCY
    c14.font = font_subtotal
    c14.alignment = Alignment(horizontal="right")
    ws_sum.cell(14, 6, "=E14/$E$21").number_format = FMT_PERCENT
    ws_sum.cell(14, 7, "Total Project Cost Base").alignment = Alignment(horizontal="center")

    # Profit Box (Row 16)
    ws_sum.cell(16, 1, "PROF_PARAM").alignment = Alignment(horizontal="center")
    ws_sum.cell(16, 2, "Contractor Profit Margin Percentage").alignment = Alignment(horizontal="left")
    ws_sum.cell(16, 3, "Master Contractor Operating Margin Parameter").alignment = Alignment(horizontal="left")
    ws_sum.cell(16, 4, "='Rates & Parameters'!$D$21").alignment = Alignment(horizontal="center")
    c16 = ws_sum.cell(16, 5, "='Rates & Parameters'!$D$21")
    c16.number_format = FMT_PERCENT
    c16.font = font_profit_pct
    c16.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    c16.border = border_profit_box
    c16.alignment = Alignment(horizontal="right")
    ws_sum.cell(16, 6, "-").alignment = Alignment(horizontal="center")
    ws_sum.cell(16, 7, "Editable Master Parameter").alignment = Alignment(horizontal="center")

    # Project Costs A (Row 17)
    ws_sum.cell(17, 1, "COST_A").alignment = Alignment(horizontal="center")
    ws_sum.cell(17, 2, "PROJECT COSTS A (EXCL. VAT)").alignment = Alignment(horizontal="left")
    ws_sum.cell(17, 3, "= Direct Costs + P&G + Contingencies").alignment = Alignment(horizontal="left")
    ws_sum.cell(17, 4, "= E14").alignment = Alignment(horizontal="center")
    c17 = ws_sum.cell(17, 5, "=E14")
    c17.number_format = FMT_CURRENCY
    c17.font = font_data_bold
    c17.alignment = Alignment(horizontal="right")
    ws_sum.cell(17, 6, "=E17/$E$21").number_format = FMT_PERCENT
    ws_sum.cell(17, 7, "Cost Base A").alignment = Alignment(horizontal="center")

    # Monthly Run-Rate (Row 18)
    ws_sum.cell(18, 1, "RUN_RATE").alignment = Alignment(horizontal="center")
    ws_sum.cell(18, 2, f"Monthly Average Run-Rate ({duration_months} Months)").alignment = Alignment(horizontal="left")
    ws_sum.cell(18, 3, f"= Project Costs A / {duration_months} Months").alignment = Alignment(horizontal="left")
    ws_sum.cell(18, 4, f"= E17 / {duration_months}").alignment = Alignment(horizontal="center")
    c18 = ws_sum.cell(18, 5, f"=E17/{duration_months}")
    c18.number_format = FMT_CURRENCY
    c18.font = font_data_bold
    c18.alignment = Alignment(horizontal="right")
    ws_sum.cell(18, 6, "-").alignment = Alignment(horizontal="center")
    ws_sum.cell(18, 7, "Monthly operational burn").alignment = Alignment(horizontal="center")

    # Contractor Profit Amount (Row 19)
    ws_sum.cell(19, 1, "PROFIT_VAL").alignment = Alignment(horizontal="center")
    ws_sum.cell(19, 2, "Contractor Net Profit Margin Value").alignment = Alignment(horizontal="left")
    ws_sum.cell(19, 3, "= Project Costs A * Profit Margin Rate").alignment = Alignment(horizontal="left")
    ws_sum.cell(19, 4, "= E17 * E16").alignment = Alignment(horizontal="center")
    c19 = ws_sum.cell(19, 5, "=E17*E16")
    c19.number_format = FMT_CURRENCY
    c19.font = font_profit_pct
    c19.alignment = Alignment(horizontal="right")
    ws_sum.cell(19, 6, "=E19/$E$21").number_format = FMT_PERCENT
    ws_sum.cell(19, 7, "Operating Return").alignment = Alignment(horizontal="center")

    # Sub-Total B Tender Price (Row 21)
    ws_sum.cell(21, 1, "SUB_B").alignment = Alignment(horizontal="center")
    ws_sum.cell(21, 2, "SUB-TOTAL (B) TENDER PRICE (EXCL. VAT)").alignment = Alignment(horizontal="left")
    ws_sum.cell(21, 3, "= Project Costs A + Contractor Net Profit").alignment = Alignment(horizontal="left")
    ws_sum.cell(21, 4, "= E17 + E19").alignment = Alignment(horizontal="center")
    c21 = ws_sum.cell(21, 5, "=E17+E19")
    c21.number_format = FMT_CURRENCY
    c21.font = font_subtotal
    c21.alignment = Alignment(horizontal="right")
    ws_sum.cell(21, 6, "=E21/$E$21").number_format = FMT_PERCENT
    ws_sum.cell(21, 7, "Net Contract Proposal").alignment = Alignment(horizontal="center")

    # SARS VAT 15% (Row 22)
    ws_sum.cell(22, 1, "VAT_15").alignment = Alignment(horizontal="center")
    ws_sum.cell(22, 2, "Value Added Tax (15% Statutory SARS VAT)").alignment = Alignment(horizontal="left")
    ws_sum.cell(22, 3, "= Subtotal B Tender Price * Statutory VAT Rate").alignment = Alignment(horizontal="left")
    ws_sum.cell(22, 4, "='Rates & Parameters'!$D$22").alignment = Alignment(horizontal="center")
    c22 = ws_sum.cell(22, 5, "=E21*'Rates & Parameters'!$D$22")
    c22.number_format = FMT_CURRENCY
    c22.font = font_data_bold
    c22.alignment = Alignment(horizontal="right")
    ws_sum.cell(22, 6, "=E22/$E$21").number_format = FMT_PERCENT
    ws_sum.cell(22, 7, "Statutory Tax").alignment = Alignment(horizontal="center")

    # TOTAL TENDER PRICE C INCL VAT (Row 23)
    ws_sum.cell(23, 1, "TOTAL_C").alignment = Alignment(horizontal="center")
    ws_sum.cell(23, 2, "TOTAL PROJECT COSTS C (TENDER PRICE INCL. VAT)").alignment = Alignment(horizontal="left")
    ws_sum.cell(23, 3, "= Subtotal B Tender Price + Value Added Tax").alignment = Alignment(horizontal="left")
    ws_sum.cell(23, 4, "= E21 + E22").alignment = Alignment(horizontal="center")
    c23 = ws_sum.cell(23, 5, "=E21+E22")
    c23.number_format = FMT_CURRENCY
    c23.font = font_grandtotal
    c23.alignment = Alignment(horizontal="right")
    ws_sum.cell(23, 6, "-").alignment = Alignment(horizontal="center")
    ws_sum.cell(23, 7, "TOTAL FINAL PROPOSAL").alignment = Alignment(horizontal="center")

    # Borders & Fills for Waterfall Rows
    for r_w in [9, 10, 11, 12, 14, 16, 17, 18, 19, 21, 22, 23]:
        ws_sum.row_dimensions[r_w].height = 22
        for col_i in range(1, 8):
            cell = ws_sum.cell(r_w, col_i)
            if r_w in [11, 14, 21]:
                cell.fill = fill_subtotal
                cell.border = border_subtotal
            elif r_w == 23:
                cell.fill = fill_grandtotal
                cell.border = border_grandtotal
                if col_i != 5:
                    cell.font = font_grandtotal
            elif r_w == 16:
                if col_i != 5:
                    cell.border = border_cell
                    cell.font = font_data_bold
            else:
                cell.border = border_cell
                if r_w % 2 == 1:
                    cell.fill = fill_zebra
                if col_i not in [5, 6]:
                    cell.font = font_data

    ws_sum.column_dimensions["A"].width = 14
    ws_sum.column_dimensions["B"].width = 46
    ws_sum.column_dimensions["C"].width = 44
    ws_sum.column_dimensions["D"].width = 24
    ws_sum.column_dimensions["E"].width = 26
    ws_sum.column_dimensions["F"].width = 14
    ws_sum.column_dimensions["G"].width = 24

    # =========================================================================
    # DEDICATED PHASE SHEETS
    # =========================================================================
    for p_idx, phase in enumerate(phases, 1):
        p_name = phase.get("name", f"Phase {p_idx}")
        # Sanitize sheet title (Excel limit 31 chars)
        sheet_title = p_name[:30].replace("/", "-").replace("\\", "-").replace("?", "").replace("*", "").replace("[", "").replace("]", "")
        ws_p = wb.create_sheet(title=sheet_title)
        ws_p.views.sheetView[0].showGridLines = True

        ws_p.merge_cells("A1:G1")
        ws_p["A1"] = f"{p_name.upper()} — WORK PACKAGE DETAIL"
        ws_p["A1"].font = font_main_title
        ws_p["A1"].fill = fill_main_title
        ws_p["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws_p.row_dimensions[1].height = 36

        ws_p.merge_cells("A2:G2")
        ws_p["A2"] = f"Surveyed Area: {phase.get('footprint_m2', 0):,.0f} m² | Structures: {phase.get('structures', 0)} | Concrete: {phase.get('concrete_m3', 0):,.0f} m³"
        ws_p["A2"].font = font_sub_title
        ws_p["A2"].fill = fill_main_title
        ws_p["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws_p.row_dimensions[2].height = 18

        for c_idx, h in enumerate(["Item Ref", "Activity Breakdown Description", "Unit", "Quantity", "Unit Rate (ZAR)", "Direct Cost (ZAR)", "Rate Source"], 1):
            cell = ws_p.cell(4, c_idx, h)
            cell.font = font_tbl_hdr
            cell.fill = fill_tbl_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_cell
        ws_p.row_dimensions[4].height = 24

        p_items = [
            (f"{p_idx}.1", "Site Establishment, mobilization & environmental setup", "Lot", 1, "='Rates & Parameters'!$D$6", "'Rates & Parameters'!D6"),
            (f"{p_idx}.2", "Cut up and dismantle structural steel, roofing & trusses", "R/struct", phase.get("structures", 2), "='Rates & Parameters'!$D$7", "'Rates & Parameters'!D7"),
            (f"{p_idx}.3", "Break up concrete slabs & plinths (±1.5m below NGL)", "m³", phase.get("concrete_m3", 500), "='Rates & Parameters'!$D$10", "'Rates & Parameters'!D10"),
            (f"{p_idx}.4", "Remove, load and haul clean building rubble to stockpile", "R/ton", phase.get("rubble_tons", round(phase.get("concrete_m3", 500) * rates_cfg.get("concrete_density", 2.40))), "='Rates & Parameters'!$D$12", "'Rates & Parameters'!D12"),
            (f"{p_idx}.5", "Shaping, levelling and cleaning off area", "m²", phase.get("footprint_m2", 10000), "='Rates & Parameters'!$D$13", "'Rates & Parameters'!D13"),
        ]

        for r_i, (ref, dsc, uom, q, r_exp, src) in enumerate(p_items, 5):
            ws_p.cell(r_i, 1, ref).alignment = Alignment(horizontal="center")
            ws_p.cell(r_i, 2, dsc).alignment = Alignment(horizontal="left")
            ws_p.cell(r_i, 3, uom).alignment = Alignment(horizontal="center")
            
            c_q = ws_p.cell(r_i, 4, q)
            c_q.number_format = FMT_QTY
            c_q.alignment = Alignment(horizontal="right")

            c_r = ws_p.cell(r_i, 5, r_exp)
            c_r.number_format = FMT_CURRENCY
            c_r.alignment = Alignment(horizontal="right")

            c_t = ws_p.cell(r_i, 6, f"=D{r_i}*E{r_i}")
            c_t.number_format = FMT_CURRENCY
            c_t.alignment = Alignment(horizontal="right")
            c_t.font = font_data_bold

            ws_p.cell(r_i, 7, src).alignment = Alignment(horizontal="center")

            for col_k in range(1, 8):
                c = ws_p.cell(r_i, col_k)
                c.border = border_cell
                if col_k not in [4, 5, 6]:
                    c.font = font_data
                if r_i % 2 == 1:
                    c.fill = fill_zebra
            ws_p.row_dimensions[r_i].height = 20

        sub_row = 10
        ws_p.merge_cells(f"A{sub_row}:E{sub_row}")
        ws_p[f"A{sub_row}"] = f"TOTAL {p_name.upper()} DIRECT EXPENDITURE"
        ws_p[f"A{sub_row}"].font = font_subtotal
        ws_p[f"A{sub_row}"].alignment = Alignment(horizontal="right", vertical="center")

        c_sub = ws_p.cell(sub_row, 6, f"=SUM(F5:F9)")
        c_sub.number_format = FMT_CURRENCY
        c_sub.font = font_subtotal
        c_sub.alignment = Alignment(horizontal="right")

        for col_k in range(1, 8):
            c = ws_p.cell(sub_row, col_k)
            c.fill = fill_subtotal
            c.border = border_subtotal
        ws_p.row_dimensions[sub_row].height = 24

        ws_p.column_dimensions["A"].width = 12
        ws_p.column_dimensions["B"].width = 54
        ws_p.column_dimensions["C"].width = 14
        ws_p.column_dimensions["D"].width = 16
        ws_p.column_dimensions["E"].width = 22
        ws_p.column_dimensions["F"].width = 24
        ws_p.column_dimensions["G"].width = 24

    # Save
    wb.save(output_path)
    print(f"SUCCESS: Universal costing workbook created: {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Build Universal Demolition & Rehabilitation Financial Model")
    parser.add_argument("--preset", choices=["westgold", "village", "tau10"], default="village", help="Project preset to build")
    parser.add_argument("--config", type=str, default=None, help="Path to custom JSON configuration file")
    parser.add_argument("--output", type=str, default=None, help="Output Excel filename")

    args = parser.parse_args()

    if args.config and os.path.exists(args.config):
        with open(args.config, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    elif args.preset == "westgold":
        cfg = PRESET_WEST_GOLD
    elif args.preset == "village":
        cfg = PRESET_SENATLA_VILLAGES
    else:
        cfg = PRESET_SENATLA_VILLAGES

    if args.output:
        cfg["output_filename"] = args.output

    out_file = build_universal_costing_workbook(cfg)
    print(f"Generated: {out_file}")


if __name__ == "__main__":
    main()
