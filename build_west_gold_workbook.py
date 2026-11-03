import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage

def build_complete_west_gold_model(output_filename="West_Gold_Plant_Demolition_Rehabilitation_Model.xlsx"):
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # -------------------------------------------------------------
    # COLOR PALETTE & STYLES (Professional Mining Industrial Executive)
    # -------------------------------------------------------------
    NAVY_DARK = "1B365D"       # Primary Headers (#1B365D)
    NAVY_MED = "2E5B88"        # Section Subheaders (#2E5B88)
    GOLD_ACCENT = "C5A059"     # Gold Branding Accent (#C5A059)
    GOLD_LIGHT = "FFF9E6"      # Subtle gold tint
    GRAY_LIGHT = "F4F6F9"      # Clean zebra striping
    GRAY_SUBTOTAL = "E9EEF4"   # Subtotal row fill
    GRAY_BORDER = "D9D9D9"     # Standard cell borders
    INPUT_YELLOW = "FFF2CC"    # User editable cell fill
    INPUT_BORDER = "D4AF37"    # Gold border for editable inputs
    PROFIT_GREEN_BORDER = "274E13" # Green border for Profit box
    PROFIT_RED_TEXT = "C00000"     # Red text for Profit %
    WHITE = "FFFFFF"

    # Fonts
    font_main_title = Font(name="Segoe UI", size=15, bold=True, color=WHITE)
    font_sub_title = Font(name="Segoe UI", size=10, italic=True, color=WHITE)
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
    fill_zebra = PatternFill(start_color=GRAY_LIGHT, end_color=GRAY_LIGHT, fill_type="solid")
    fill_subtotal = PatternFill(start_color=GRAY_SUBTOTAL, end_color=GRAY_SUBTOTAL, fill_type="solid")
    fill_grandtotal = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
    fill_input = PatternFill(start_color=INPUT_YELLOW, end_color=INPUT_YELLOW, fill_type="solid")
    fill_kpi = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_card = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

    # Borders
    thin_border_side = Side(style='thin', color='D9D9D9')
    med_border_side = Side(style='medium', color='1B365D')
    double_bottom_side = Side(style='double', color='1B365D')
    gold_border_side = Side(style='thin', color=INPUT_BORDER)
    profit_border_side = Side(style='medium', color=PROFIT_GREEN_BORDER)

    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_subtotal = Border(left=thin_border_side, right=thin_border_side, top=Side(style='thin', color='1B365D'), bottom=double_bottom_side)
    border_grandtotal = Border(left=med_border_side, right=med_border_side, top=Side(style='thin', color=WHITE), bottom=Side(style='double', color=WHITE))
    border_input = Border(left=gold_border_side, right=gold_border_side, top=gold_border_side, bottom=gold_border_side)
    border_profit_box = Border(left=profit_border_side, right=profit_border_side, top=profit_border_side, bottom=profit_border_side)
    border_kpi = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # Number Formats
    FMT_CURRENCY = 'R #,##0.00'
    FMT_QTY = '#,##0'
    FMT_QTY_DEC = '#,##0.00'
    FMT_PERCENT = '0.0%'

    # Alignments
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    align_title = Alignment(horizontal="center", vertical="center")

    # =========================================================================
    # SHEET 1: RATES & PARAMETERS (MASTER CONTROL REPOSITORY)
    # =========================================================================
    ws_rates = wb.create_sheet(title="Rates & Parameters")
    ws_rates.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_rates.merge_cells("B2:G2")
    ws_rates["B2"] = "WEST GOLD PLANT - MASTER UNIT RATES & COMMERCIAL PARAMETERS"
    ws_rates["B2"].font = font_main_title
    ws_rates["B2"].fill = fill_main_title
    ws_rates["B2"].alignment = align_title
    ws_rates.row_dimensions[2].height = 34

    ws_rates.merge_cells("B3:G3")
    ws_rates["B3"] = "CENTRAL RATE REPOSITORY: Edit any rate in the yellow cells below. All 7 Phase sheets, Master BOQ, Summaries and Tender Price update automatically."
    ws_rates["B3"].font = font_sub_title
    ws_rates["B3"].fill = fill_sec_hdr
    ws_rates["B3"].alignment = align_center
    ws_rates.row_dimensions[3].height = 20

    # Section 1: Core Activity Unit Rates
    ws_rates.merge_cells("B5:G5")
    ws_rates["B5"] = "1. CORE DEMOLITION & REHABILITATION UNIT RATES (APPLIED ACROSS ALL 7 PHASES)"
    ws_rates["B5"].font = font_sec_hdr
    ws_rates["B5"].fill = fill_sec_hdr
    ws_rates["B5"].alignment = align_left
    ws_rates.row_dimensions[5].height = 24

    headers_rates = ["Item Code", "Activity Category", "Activity & Line Item Description", "UOM", "Master Unit Rate (R)", "Estimating Basis & Document Grounding"]
    for col_idx, h in enumerate(headers_rates, start=2):
        cell = ws_rates.cell(row=6, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [2, 5, 6] else align_left
        cell.border = border_cell
    ws_rates.row_dimensions[6].height = 25

    # Core rates data: Row 7 to 16
    rates_data = [
        ("R_EST", "Site Establishment", "Phase Site Establishment, Mobilization, Health & Safety, Decontamination Setup", "Lot", 35000.00, "Matches contractual screenshot rate (R 35 000,00 / Lot)"),
        ("R_STL_STD", "1. Steel & Structures", "Dismantling, cutting up & removal of standard structural steel / smelt house framing", "R/structure", 21000.00, "Matches contractual screenshot rate (R 21 000,00 / structure)"),
        ("R_STL_HVY", "1. Steel & Structures", "Heavy industrial structural steel dismantling (Mill section gantry, conveyor towers)", "R/structure", 25000.00, "Heavy plant structure allowance (higher elevation & heavy rigging)"),
        ("R_STL_WKP", "1. Steel & Structures", "Light workshop portal frames, change house trusses, roof sheeting & ties", "R/structure", 19500.00, "Matches contractual screenshot rate (R 19 500,00 / structure)"),
        ("R_CON_M3", "2. Concrete Demolition", "Breakup & stockpile reinforced concrete & brickwork above and ±1.5m below NGL", "m³", 240.00, "Heavy excavator hydraulic hammering & mechanical pulverizing per m³"),
        ("R_CON_TON", "2. Concrete Demolition", "Concrete Breakup & Stockpile on tonnage basis (equivalent @ 2.40 t/m³ density)", "R/ton", 100.00, "Tonnage equivalent rate for concrete processing"),
        ("R_HAUL_TON", "3. Loading & Hauling", "Loading and hauling of clean building rubble / pulverized material to designated dumping", "R/ton", 120.00, "Matches contractual screenshot rate exactly (R 120,00 / ton)"),
        ("R_SHAPE_M2", "4. Shaping & Levelling", "Final shaping, levelling, grading, compaction and cleaning off rehabilitated area", "R/m²", 18.50, "Dozer / grader shaping and surface clearance per m² footprint"),
        ("R_DAM_BKF", "Specialist Civil Works", "Empty process dam, dewater, backfill & compact certified fill to ±1.5m below NGL", "m³", 85.00, "Phase 3 specific scope item: Dam dewatering and engineered backfilling"),
        ("R_SILO_DEM", "Specialist Concrete", "High-elevation concrete silo demolition, specialized shearing & foundation breakup", "m³", 260.00, "Phase 4 specific scope item: Tall concrete silos & heavy raft foundation")
    ]

    for idx, r in enumerate(rates_data, start=7):
        ws_rates.cell(row=idx, column=2, value=r[0]).alignment = align_center
        ws_rates.cell(row=idx, column=3, value=r[1]).alignment = align_left
        ws_rates.cell(row=idx, column=4, value=r[2]).alignment = align_left_wrap
        ws_rates.cell(row=idx, column=5, value=r[3]).alignment = align_center
        
        rate_cell = ws_rates.cell(row=idx, column=6, value=r[4])
        rate_cell.alignment = align_right
        rate_cell.font = font_input
        rate_cell.fill = fill_input
        rate_cell.border = border_input
        rate_cell.number_format = FMT_CURRENCY
        
        ws_rates.cell(row=idx, column=7, value=r[5]).alignment = align_left_wrap
        for c in [2, 3, 4, 5, 7]:
            ws_rates.cell(row=idx, column=c).font = font_data
            ws_rates.cell(row=idx, column=c).border = border_cell
        ws_rates.row_dimensions[idx].height = 22

    # Section 2: Commercial Markups & Financial Model Parameters (Row 18 to 23)
    ws_rates.merge_cells("B18:G18")
    ws_rates["B18"] = "2. CONTRACT COMMERCIAL MARKUPS & FINANCIAL MODEL PARAMETERS"
    ws_rates["B18"].font = font_sec_hdr
    ws_rates["B18"].fill = fill_sec_hdr
    ws_rates["B18"].alignment = align_left
    ws_rates.row_dimensions[18].height = 24

    comm_headers = ["Parameter Code", "Commercial Parameter Name", "Contractual Description / Basis", "UOM", "Master Parameter Value", "Notes & Screenshot Reference"]
    for col_idx, h in enumerate(comm_headers, start=2):
        cell = ws_rates.cell(row=19, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [2, 5, 6] else align_left
        cell.border = border_cell
    ws_rates.row_dimensions[19].height = 25

    comm_params = [
        ("P_PG", "Preliminaries & General (P&G)", "Site overheads, health & safety, management, insurances & environmental supervision", "%", 0.10, "10.0% P&G (Matches Screenshot 2 row: 'Preliminaries and General (10%)')"),
        ("P_CONT", "Contingencies Allowance", "Contingency reserve for hidden foundations, unknown sub-surface contamination & weather delays", "%", 0.10, "10.0% Contingencies (Matches Screenshot 2 row: 'Contingencies (10%)')"),
        ("P_PROFIT", "Contractor Profit Margin", "Target contractor net operating margin applied to Project Costs A", "%", 0.05, "5.0% Profit (Matches Screenshot 3 highlighted box: 'Profit Percentage 5%')"),
        ("P_VAT", "Value Added Tax (VAT)", "South African Revenue Service (SARS) standard statutory VAT rate", "%", 0.15, "15.0% VAT (Matches Screenshot 3 row: '15% R 396 007,23')")
    ]

    for idx, p in enumerate(comm_params, start=20):
        ws_rates.cell(row=idx, column=2, value=p[0]).alignment = align_center
        ws_rates.cell(row=idx, column=3, value=p[1]).alignment = align_left
        ws_rates.cell(row=idx, column=4, value=p[2]).alignment = align_left_wrap
        ws_rates.cell(row=idx, column=5, value=p[3]).alignment = align_center
        
        param_cell = ws_rates.cell(row=idx, column=6, value=p[4])
        param_cell.alignment = align_right
        param_cell.font = font_input
        param_cell.fill = fill_input
        param_cell.border = border_input
        param_cell.number_format = FMT_PERCENT
        
        ws_rates.cell(row=idx, column=7, value=p[5]).alignment = align_left_wrap
        for c in [2, 3, 4, 5, 7]:
            ws_rates.cell(row=idx, column=c).font = font_data
            ws_rates.cell(row=idx, column=c).border = border_cell
        ws_rates.row_dimensions[idx].height = 22

    # Section 3: Technical Assumptions & Conversion Factors (Row 25 to 31)
    ws_rates.merge_cells("B25:G25")
    ws_rates["B25"] = "3. TECHNICAL ENGINEERING ASSUMPTIONS & MATERIAL CONVERSION FACTORS"
    ws_rates["B25"].font = font_sec_hdr
    ws_rates["B25"].fill = fill_sec_hdr
    ws_rates["B25"].alignment = align_left
    ws_rates.row_dimensions[25].height = 24

    tech_headers = ["Factor Code", "Technical Parameter", "Engineering Scope & Physical Definition", "Unit", "Factor Value", "Conversion Application"]
    for col_idx, h in enumerate(tech_headers, start=2):
        cell = ws_rates.cell(row=26, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [2, 5, 6] else align_left
        cell.border = border_cell
    ws_rates.row_dimensions[26].height = 25

    tech_params = [
        ("F_DENSITY", "Reinforced Concrete Density", "Bulk density of reinforced industrial concrete foundations & slabs", "ton/m³", 2.40, "Converts concrete volume (m³) to haulage tonnage (ton)"),
        ("F_BULK", "Bulking Swell Factor", "Mechanical fragmentation bulking ratio for loose loading and haulage", "ratio", 1.30, "Loose volume check for tipper truck payload calculation"),
        ("F_DEPTH", "Below-Ground Demolition Depth", "Standard rehabilitation clearance depth below natural ground level (NGL)", "meters", 1.50, "Contract requirement: ±1.5m below NGL in every phase"),
        ("P_TOTAL_AREA", "Total Plant Footprint Area", "Sum of Phase 1 through Phase 7 surveyed aerial surface areas", "m²", 93072.00, "Aerial survey boundary measurement sum"),
        ("P_DURATION", "Total Contract Project Duration", "Estimated project duration spanning all 7 operational phases", "Months", 22.00, "Contract schedule: 4+3+4+4+2+2+3 = 22 Months")
    ]

    for idx, t in enumerate(tech_params, start=27):
        ws_rates.cell(row=idx, column=2, value=t[0]).alignment = align_center
        ws_rates.cell(row=idx, column=3, value=t[1]).alignment = align_left
        ws_rates.cell(row=idx, column=4, value=t[2]).alignment = align_left_wrap
        ws_rates.cell(row=idx, column=5, value=t[3]).alignment = align_center
        
        val_cell = ws_rates.cell(row=idx, column=6, value=t[4])
        val_cell.alignment = align_right
        val_cell.font = font_input
        val_cell.fill = fill_input
        val_cell.border = border_input
        val_cell.number_format = FMT_QTY_DEC if t[3] in ["ton/m³", "ratio", "meters"] else FMT_QTY
        
        ws_rates.cell(row=idx, column=7, value=t[5]).alignment = align_left_wrap
        for c in [2, 3, 4, 5, 7]:
            ws_rates.cell(row=idx, column=c).font = font_data
            ws_rates.cell(row=idx, column=c).border = border_cell
        ws_rates.row_dimensions[idx].height = 22

    # Column widths for Rates & Parameters
    ws_rates.column_dimensions["A"].width = 3
    ws_rates.column_dimensions["B"].width = 16
    ws_rates.column_dimensions["C"].width = 28
    ws_rates.column_dimensions["D"].width = 54
    ws_rates.column_dimensions["E"].width = 14
    ws_rates.column_dimensions["F"].width = 22
    ws_rates.column_dimensions["G"].width = 46

    # =========================================================================
    # SHEET 2: MASTER BOQ BREAKDOWN (CONSOLIDATED SCHEDULE AS PER SCREENSHOT 1)
    # =========================================================================
    ws_boq = wb.create_sheet(title="Master BOQ Breakdown")
    ws_boq.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_boq.merge_cells("A2:G2")
    ws_boq["A2"] = "WEST GOLD PLANT - CONSOLIDATED BILL OF QUANTITIES"
    ws_boq["A2"].font = font_main_title
    ws_boq["A2"].fill = fill_main_title
    ws_boq["A2"].alignment = align_title
    ws_boq.row_dimensions[2].height = 32

    ws_boq.merge_cells("A3:G3")
    ws_boq["A3"] = "Demolition, Concrete Breakup, Load and Haul Rubble & Surface Rehabilitation (Phase 01 to Phase 07)"
    ws_boq["A3"].font = font_sub_title
    ws_boq["A3"].fill = fill_sec_hdr
    ws_boq["A3"].alignment = align_center
    ws_boq.row_dimensions[3].height = 20

    boq_headers = ["Item", "Plant Phase / Section", "Description", "UOM", "Qty", "Unit Rate (R)", "Total (R)"]
    for col_idx, h in enumerate(boq_headers, start=1):
        cell = ws_boq.cell(row=5, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [1, 4, 5] else (align_right if col_idx in [6, 7] else align_left)
        cell.border = border_cell
    ws_boq.row_dimensions[5].height = 26

    # Data structures for 7 phases
    phases_master = [
        (
            1, "Phase 01: Smelt House Area", 10574.00, 4,
            [
                ("Site Establishment, mobilization, health & safety and decontamination setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Cut up and remove all reinforcing material / steel structures of smelt house (R/structure)", "R/structure", 3, "='Rates & Parameters'!$F$8"),
                ("2. Breakup & stockpile concrete and brickwork above and ±1.5m below NGL (m³)", "m³", 1250, "='Rates & Parameters'!$F$11"),
                ("3. Loading and hauling of clean building rubble to designated dumping (R/ton) [3,000 t]", "R/ton", 3000, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (10,574 m²)", "m²", 10574, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            2, "Phase 02: Thickener Dams Area", 10890.00, 3,
            [
                ("Site Establishment, mobilization, health & safety and environmental setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Cut up and remove all reinforcing material / steel structures & rake launders (R/structure)", "R/structure", 3, "='Rates & Parameters'!$F$8"),
                ("2. Breakup & stockpile concrete & brickwork above and ±1.5m below NGL – Thickener Dams", "m³", 1600, "='Rates & Parameters'!$F$11"),
                ("3. Loading and hauling of clean building rubble to designated dumping (R/ton) [3,840 t]", "R/ton", 3840, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (10,890 m²)", "m²", 10890, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            3, "Phase 03: Mill Section & Dams Area", 16823.00, 4,
            [
                ("Site Establishment, mobilization, heavy crane staging & safety setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Dismantling of mill section heavy steel structures, gantry frames & reinforcing (R/structure)", "R/structure", 6, "='Rates & Parameters'!$F$9"),
                ("2. Breakup & stockpile all concrete and brickwork above and ±1.5m below NGL incl. mill plinths", "m³", 2800, "='Rates & Parameters'!$F$11"),
                ("   Specialist Civil: Empty process dams and backfill with materials up ±1.5m below NGL", "m³", 3500, "='Rates & Parameters'!$F$15"),
                ("3. Loading and hauling of clean pulverized rubble to designated dumping (R/ton) [6,720 t]", "R/ton", 6720, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (16,823 m²)", "m²", 16823, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            4, "Phase 04: Silos & Foundations Area", 12763.00, 4,
            [
                ("Site Establishment, high-elevation safety netting & mobilization setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Dismantling of steel structures, conveyor headframes & high chutes (R/structure)", "R/structure", 5, "='Rates & Parameters'!$F$8"),
                ("2. Demolition of concrete silos and breakup & stockpile foundations above & ±1.5m below NGL", "m³", 3200, "='Rates & Parameters'!$F$11"),
                ("3. Loading and hauling of clean pulverized material to designated dumping (R/ton) [7,680 t]", "R/ton", 7680, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (12,763 m²)", "m²", 12763, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            5, "Phase 05: Workshops Area", 10073.00, 2,
            [
                ("Site Establishment, workshop isolation & mobilization setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Dismantling of workshops steel structures, portal frames & overhead cranes (R/structure)", "R/structure", 4, "='Rates & Parameters'!$F$10"),
                ("2. Breakup & stockpile concrete slabs, footings & aprons above & ±1.5m below NGL (m³)", "m³", 1300, "='Rates & Parameters'!$F$11"),
                ("3. Loading and hauling of clean pulverized material to designated dumping (R/ton) [3,120 t]", "R/ton", 3120, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (10,073 m²)", "m²", 10073, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            6, "Phase 06: Conveyor Structures Area", 13607.00, 2,
            [
                ("Site Establishment, overland corridor access & Eskom boundary safety setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Dismantling of conveyor structure steel structures, gantry frames & trestles (R/structure)", "R/structure", 6, "='Rates & Parameters'!$F$8"),
                ("2. Breaking concrete columns, counterweights and plinths above and ±1.5m below NGL (m³)", "m³", 1150, "='Rates & Parameters'!$F$11"),
                ("3. Loading and hauling of clean pulverized material to designated dumping (R/ton) [2,760 t]", "R/ton", 2760, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (13,607 m²)", "m²", 13607, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            7, "Phase 07: Main Gate & Change Houses Area", 18342.00, 3,
            [
                ("Site Establishment, main gate traffic control & mobilization setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Breakup all workshops and change houses structures from the main gate (R/structure)", "R/structure", 8, "='Rates & Parameters'!$F$10"),
                ("2. Breakup & stockpile concrete and brickwork masonry above and ±1.5m below NGL (m³)", "m³", 2100, "='Rates & Parameters'!$F$11"),
                ("3. Loading and hauling of clean pulverized material to designated dumping (R/ton) [5,040 t]", "R/ton", 5040, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (18,342 m²)", "m²", 18342, "='Rates & Parameters'!$F$14")
            ]
        )
    ]

    curr_row = 6
    phase_subtotal_rows = {}

    for p in phases_master:
        item_num, phase_title, area_m2, duration_mo, items = p
        start_row = curr_row
        num_items = len(items)

        end_phase_row = start_row + num_items - 1
        ws_boq.merge_cells(start_row=start_row, start_column=1, end_row=end_phase_row, end_column=1)
        ws_boq.cell(row=start_row, column=1, value=item_num).alignment = align_center
        ws_boq.cell(row=start_row, column=1).font = font_data_bold

        ws_boq.merge_cells(start_row=start_row, start_column=2, end_row=end_phase_row, end_column=2)
        cell_phase_title = ws_boq.cell(row=start_row, column=2, value=f"{phase_title}\n(Area: {area_m2:,.2f} m² | ±{duration_mo} Mos)")
        cell_phase_title.alignment = align_center
        cell_phase_title.font = font_data_bold

        for itm_idx, itm in enumerate(items):
            r_idx = start_row + itm_idx
            desc, uom, qty, rate_ref = itm

            ws_boq.cell(row=r_idx, column=3, value=desc).alignment = align_left
            ws_boq.cell(row=r_idx, column=3).font = font_data
            
            ws_boq.cell(row=r_idx, column=4, value=uom).alignment = align_center
            ws_boq.cell(row=r_idx, column=4).font = font_data
            
            c_qty = ws_boq.cell(row=r_idx, column=5, value=qty)
            c_qty.alignment = align_right
            c_qty.font = font_data
            c_qty.number_format = FMT_QTY if isinstance(qty, int) else FMT_QTY_DEC
            
            c_rate = ws_boq.cell(row=r_idx, column=6, value=rate_ref)
            c_rate.alignment = align_right
            c_rate.font = font_data
            c_rate.number_format = FMT_CURRENCY
            
            c_tot = ws_boq.cell(row=r_idx, column=7, value=f"=E{r_idx}*F{r_idx}")
            c_tot.alignment = align_right
            c_tot.font = font_data_bold
            c_tot.number_format = FMT_CURRENCY

            for c in range(1, 8):
                ws_boq.cell(row=r_idx, column=c).border = border_cell
            ws_boq.row_dimensions[r_idx].height = 22

        # Sub-Total Row
        sub_row = end_phase_row + 1
        phase_subtotal_rows[item_num] = sub_row
        ws_boq.merge_cells(start_row=sub_row, start_column=1, end_row=sub_row, end_column=5)
        ws_boq.cell(row=sub_row, column=1, value=f"Sub-Total - {phase_title}").alignment = align_right
        ws_boq.cell(row=sub_row, column=1).font = font_subtotal
        
        ws_boq.cell(row=sub_row, column=6, value="Sub-Total").alignment = align_right
        ws_boq.cell(row=sub_row, column=6).font = font_subtotal
        
        c_subtot = ws_boq.cell(row=sub_row, column=7, value=f"=SUM(G{start_row}:G{end_phase_row})")
        c_subtot.alignment = align_right
        c_subtot.font = font_subtotal
        c_subtot.number_format = FMT_CURRENCY

        for c in range(1, 8):
            cell = ws_boq.cell(row=sub_row, column=c)
            cell.fill = fill_subtotal
            cell.border = border_subtotal
        ws_boq.row_dimensions[sub_row].height = 24

        curr_row = sub_row + 1

    # Grand Total Row for Master BOQ
    row_gt = curr_row
    ws_boq.merge_cells(start_row=row_gt, start_column=1, end_row=row_gt, end_column=6)
    ws_boq.cell(row=row_gt, column=1, value="TOTAL DIRECT PROJECT COST (EXCLUDING P&G, CONTINGENCIES & PROFIT)").alignment = align_right
    ws_boq.cell(row=row_gt, column=1).font = font_grandtotal

    sum_terms = ",".join([f"G{phase_subtotal_rows[k]}" for k in sorted(phase_subtotal_rows.keys())])
    c_gt = ws_boq.cell(row=row_gt, column=7, value=f"=SUM({sum_terms})")
    c_gt.alignment = align_right
    c_gt.font = font_grandtotal
    c_gt.number_format = FMT_CURRENCY

    for c in range(1, 8):
        cell = ws_boq.cell(row=row_gt, column=c)
        cell.fill = fill_grandtotal
        cell.border = border_grandtotal
    ws_boq.row_dimensions[row_gt].height = 26

    # Column widths for Master BOQ
    ws_boq.column_dimensions["A"].width = 8
    ws_boq.column_dimensions["B"].width = 28
    ws_boq.column_dimensions["C"].width = 52
    ws_boq.column_dimensions["D"].width = 14
    ws_boq.column_dimensions["E"].width = 14
    ws_boq.column_dimensions["F"].width = 18
    ws_boq.column_dimensions["G"].width = 20

    # =========================================================================
    # SHEET 3: EXECUTIVE SUMMARY & FINANCIALS (SCREENSHOT 2 & 3 REPLICATION)
    # =========================================================================
    ws_sum = wb.create_sheet(title="Executive Summary & Financials")
    ws_sum.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_sum.merge_cells("B2:I2")
    ws_sum["B2"] = "WEST GOLD PLANT - DEMOLITION AND REHABILITATION PROJECT"
    ws_sum["B2"].font = font_main_title
    ws_sum["B2"].fill = fill_main_title
    ws_sum["B2"].alignment = align_title
    ws_sum.row_dimensions[2].height = 34

    ws_sum.merge_cells("B3:I3")
    ws_sum["B3"] = "EXECUTIVE SUMMARY, COST WATERFALL, COMMERCIAL PROFIT MODEL & TENDER PRICING"
    ws_sum["B3"].font = font_sub_title
    ws_sum["B3"].fill = fill_sec_hdr
    ws_sum["B3"].alignment = align_center
    ws_sum.row_dimensions[3].height = 20

    # KPI Summary Cards Block
    for title, formula, unit, merge_val, merge_lbl in [
        ("TOTAL SITE AREA", "='Rates & Parameters'!$F$30", " m²", "B5:C5", "B6:C6"),
        ("EST. DURATION", "='Rates & Parameters'!$F$31", " Months", "D5:E5", "D6:E6"),
        ("TOTAL RUBBLE TO HAUL", 32160, " Tonnes", "F5:G5", "F6:G6"),
        ("TOTAL STRUCTURES", 35, " Structures", "H5:I5", "H6:I6")
    ]:
        ws_sum.merge_cells(merge_val)
        top_left = merge_val.split(":")[0]
        cell_v = ws_sum[top_left]
        cell_v.value = formula
        cell_v.font = font_kpi_num
        cell_v.alignment = align_center
        cell_v.fill = fill_kpi
        if isinstance(formula, str) and "Rates" in formula:
            cell_v.number_format = '#,##0.00" m²"' if "30" in formula else '#,##0" Months"'
        else:
            cell_v.number_format = f'#,##0"{unit}"'

        ws_sum.merge_cells(merge_lbl)
        lbl_top = merge_lbl.split(":")[0]
        cell_l = ws_sum[lbl_top]
        cell_l.value = title
        cell_l.font = font_kpi_lbl
        cell_l.alignment = align_center
        cell_l.fill = fill_kpi

        for row in ws_sum[merge_val]:
            for c in row:
                c.border = border_kpi
        for row in ws_sum[merge_lbl]:
            for c in row:
                c.border = border_kpi

    ws_sum.row_dimensions[5].height = 24
    ws_sum.row_dimensions[6].height = 18

    # Table 1: Summary of Demolition, Load and Hauling of Rubble (Matches Screenshot 2)
    row_t1_hdr = 8
    ws_sum.merge_cells(f"B{row_t1_hdr}:I{row_t1_hdr}")
    ws_sum[f"B{row_t1_hdr}"] = "Summary of Demolition, Load and Hauling of Rubble"
    ws_sum[f"B{row_t1_hdr}"].font = font_sec_hdr
    ws_sum[f"B{row_t1_hdr}"].fill = fill_sec_hdr
    ws_sum[f"B{row_t1_hdr}"].alignment = align_center
    ws_sum.row_dimensions[row_t1_hdr].height = 24

    sum_col_headers = ["Item", "Phase / Plant Section", "Scope Summary", "Footprint Area (m²)", "Duration (Months)", "Structures Qty", "Rubble (Tonnes)", "Amount (R)"]
    for col_idx, h in enumerate(sum_col_headers, start=2):
        cell = ws_sum.cell(row=row_t1_hdr+1, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [2, 5, 6, 7, 8] else (align_right if col_idx == 9 else align_left)
        cell.border = border_cell
    ws_sum.row_dimensions[row_t1_hdr+1].height = 25

    sum_phase_rows = [
        (1, "Phase 01: Smelt House Area", "Smelt house steel dismantling, concrete breakup ±1.5m below NGL, load & haul rubble, shape & level", 10574.00, 4, 3, 3000, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[1]}"),
        (2, "Phase 02: Thickener Dams Area", "Thickener dams concrete breakup ±1.5m below NGL, steel launders, load & haul, shape & level", 10890.00, 3, 3, 3840, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[2]}"),
        (3, "Phase 03: Mill Section & Dams Area", "Mill heavy steel structures, mill concrete plinths, empty & backfill process dams, load & haul", 16823.00, 4, 6, 6720, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[3]}"),
        (4, "Phase 04: Silos & Foundations Area", "High concrete silos demolition, silo foundations breakup ±1.5m below NGL, load & haul", 12763.00, 4, 5, 7680, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[4]}"),
        (5, "Phase 05: Workshops Area", "Workshop steel portal frames, overhead cranes, concrete slabs breakup, load & haul", 10073.00, 2, 4, 3120, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[5]}"),
        (6, "Phase 06: Conveyor Structures Area", "Conveyor overland gantries, concrete trestle columns, load & haul, shape & level", 13607.00, 2, 6, 2760, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[6]}"),
        (7, "Phase 07: Main Gate & Change Houses Area", "Main gate workshops, change houses, brick masonry breakup, load & haul, shape & level", 18342.00, 3, 8, 5040, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[7]}")
    ]

    for idx, p in enumerate(sum_phase_rows, start=row_t1_hdr+2):
        ws_sum.cell(row=idx, column=2, value=p[0]).alignment = align_center
        ws_sum.cell(row=idx, column=3, value=p[1]).alignment = align_left
        ws_sum.cell(row=idx, column=4, value=p[2]).alignment = align_left_wrap
        
        c_a = ws_sum.cell(row=idx, column=5, value=p[3])
        c_a.alignment = align_right
        c_a.number_format = FMT_QTY_DEC
        
        c_d = ws_sum.cell(row=idx, column=6, value=p[4])
        c_d.alignment = align_center
        c_d.number_format = FMT_QTY
        
        c_s = ws_sum.cell(row=idx, column=7, value=p[5])
        c_s.alignment = align_center
        c_s.number_format = FMT_QTY

        c_r = ws_sum.cell(row=idx, column=8, value=p[6])
        c_r.alignment = align_right
        c_r.number_format = FMT_QTY
        
        c_amt = ws_sum.cell(row=idx, column=9, value=p[7])
        c_amt.alignment = align_right
        c_amt.font = font_data_bold
        c_amt.number_format = FMT_CURRENCY

        fill_row = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)
        for c in range(2, 10):
            ws_sum.cell(row=idx, column=c).border = border_cell
            if fill_row.fill_type:
                ws_sum.cell(row=idx, column=c).fill = fill_row
        ws_sum.row_dimensions[idx].height = 23

    # Financial Waterfall Rows (Modeled EXACTLY on Screenshot 2)
    # Row 17: Sub-Total (Excluding Preliminaries and General)
    r_sub_excl_pg = 17
    ws_sum.merge_cells(f"B{r_sub_excl_pg}:H{r_sub_excl_pg}")
    ws_sum.cell(row=r_sub_excl_pg, column=2, value="Sub-Total (Excluding Preliminaries and General)").alignment = align_right
    ws_sum.cell(row=r_sub_excl_pg, column=2).font = font_subtotal

    c_sub_val = ws_sum.cell(row=r_sub_excl_pg, column=9, value="=SUM(I10:I16)")
    c_sub_val.alignment = align_right
    c_sub_val.font = font_subtotal
    c_sub_val.number_format = FMT_CURRENCY

    for c in range(2, 10):
        ws_sum.cell(row=r_sub_excl_pg, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_sub_excl_pg, column=c).border = border_subtotal
    ws_sum.row_dimensions[r_sub_excl_pg].height = 24

    # Row 18: Preliminaries and General (10%)
    r_pg = 18
    ws_sum.merge_cells(f"B{r_pg}:H{r_pg}")
    ws_sum.cell(row=r_pg, column=2, value="Preliminaries and General (10%)").alignment = align_right
    ws_sum.cell(row=r_pg, column=2).font = font_data_bold
    c_pg_val = ws_sum.cell(row=r_pg, column=9, value="=I17*'Rates & Parameters'!$F$20")
    c_pg_val.alignment = align_right
    c_pg_val.font = font_data_bold
    c_pg_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_pg, column=c).border = border_cell
    ws_sum.row_dimensions[r_pg].height = 22

    # Row 19: Sub-Total (Excluding Contingencies)
    r_sub_excl_cont = 19
    ws_sum.merge_cells(f"B{r_sub_excl_cont}:H{r_sub_excl_cont}")
    ws_sum.cell(row=r_sub_excl_cont, column=2, value="Sub-Total (Excluding Contingencies)").alignment = align_right
    ws_sum.cell(row=r_sub_excl_cont, column=2).font = font_subtotal
    c_sub_excl_cont_val = ws_sum.cell(row=r_sub_excl_cont, column=9, value="=I17+I18")
    c_sub_excl_cont_val.alignment = align_right
    c_sub_excl_cont_val.font = font_subtotal
    c_sub_excl_cont_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_sub_excl_cont, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_sub_excl_cont, column=c).border = border_cell
    ws_sum.row_dimensions[r_sub_excl_cont].height = 23

    # Row 20: Contingencies (10%)
    r_cont = 20
    ws_sum.merge_cells(f"B{r_cont}:H{r_cont}")
    ws_sum.cell(row=r_cont, column=2, value="Contingencies (10%)").alignment = align_right
    ws_sum.cell(row=r_cont, column=2).font = font_data_bold
    c_cont_val = ws_sum.cell(row=r_cont, column=9, value="=I19*'Rates & Parameters'!$F$21")
    c_cont_val.alignment = align_right
    c_cont_val.font = font_data_bold
    c_cont_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_cont, column=c).border = border_cell
    ws_sum.row_dimensions[r_cont].height = 22

    # Row 21: Sub Total: Inclusive of Preliminaries and General & Contingencies
    r_sub_incl_all = 21
    ws_sum.merge_cells(f"B{r_sub_incl_all}:H{r_sub_incl_all}")
    ws_sum.cell(row=r_sub_incl_all, column=2, value="Sub Total: Inclusive of Preliminaries and General & Contingencies").alignment = align_right
    ws_sum.cell(row=r_sub_incl_all, column=2).font = font_subtotal
    c_sub_incl_val = ws_sum.cell(row=r_sub_incl_all, column=9, value="=I19+I20")
    c_sub_incl_val.alignment = align_right
    c_sub_incl_val.font = font_subtotal
    c_sub_incl_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_sub_incl_all, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_sub_incl_all, column=c).border = border_subtotal
    ws_sum.row_dimensions[r_sub_incl_all].height = 24

    # Table 2: Financial Model & Profit Calculation (Matches Screenshot 3 EXACTLY)
    # Header line with Profit Percentage
    r_prof_hdr = 23
    ws_sum.merge_cells(f"G{r_prof_hdr}:H{r_prof_hdr}")
    ws_sum[f"G{r_prof_hdr}"] = "Profit Percentage"
    ws_sum[f"G{r_prof_hdr}"].font = font_data_bold
    ws_sum[f"G{r_prof_hdr}"].alignment = align_right
    ws_sum[f"G{r_prof_hdr}"].border = border_cell

    c_prof_pct = ws_sum.cell(row=r_prof_hdr, column=9, value="='Rates & Parameters'!$F$22")
    c_prof_pct.font = font_profit_pct
    c_prof_pct.alignment = align_center
    c_prof_pct.number_format = FMT_PERCENT
    c_prof_pct.border = border_profit_box
    ws_sum.row_dimensions[r_prof_hdr].height = 26

    # Sub-header: Monthly Average
    r_monthly = 24
    ws_sum.merge_cells(f"G{r_monthly}:H{r_monthly}")
    ws_sum[f"G{r_monthly}"] = "Monthly"
    ws_sum[f"G{r_monthly}"].font = font_data_bold
    ws_sum[f"G{r_monthly}"].alignment = align_right
    ws_sum[f"G{r_monthly}"].border = border_cell

    c_monthly_val = ws_sum.cell(row=r_monthly, column=9, value="=I21/'Rates & Parameters'!$F$31")
    c_monthly_val.font = font_data
    c_monthly_val.alignment = align_right
    c_monthly_val.number_format = FMT_CURRENCY
    c_monthly_val.border = border_cell
    ws_sum.row_dimensions[r_monthly].height = 22

    # Row Profit
    r_profit = 25
    ws_sum.merge_cells(f"B{r_profit}:H{r_profit}")
    ws_sum[f"B{r_profit}"] = "Profit"
    ws_sum[f"B{r_profit}"].font = font_data_bold
    ws_sum[f"B{r_profit}"].alignment = align_right
    c_prof_val = ws_sum.cell(row=r_profit, column=9, value="=I26*I23")
    c_prof_val.font = font_data_bold
    c_prof_val.alignment = align_right
    c_prof_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_profit, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_profit, column=c).border = border_cell
    ws_sum.row_dimensions[r_profit].height = 22

    # Row Project costs A (Excl. VAT)
    r_proj_cost = 26
    ws_sum.merge_cells(f"B{r_proj_cost}:H{r_proj_cost}")
    ws_sum[f"B{r_proj_cost}"] = "Project costs A (Excl. VAT)"
    ws_sum[f"B{r_proj_cost}"].font = font_data_bold
    ws_sum[f"B{r_proj_cost}"].alignment = align_right
    c_pc_val = ws_sum.cell(row=r_proj_cost, column=9, value="=I21")
    c_pc_val.font = font_data_bold
    c_pc_val.alignment = align_right
    c_pc_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_proj_cost, column=c).border = border_cell
    ws_sum.row_dimensions[r_proj_cost].height = 22

    # Row Sub-Total
    r_subtot_fin = 27
    ws_sum.merge_cells(f"B{r_subtot_fin}:H{r_subtot_fin}")
    ws_sum[f"B{r_subtot_fin}"] = "Sub-Total"
    ws_sum[f"B{r_subtot_fin}"].font = font_subtotal
    ws_sum[f"B{r_subtot_fin}"].alignment = align_right
    c_st_fin = ws_sum.cell(row=r_subtot_fin, column=9, value="=I26+I25")
    c_st_fin.font = font_subtotal
    c_st_fin.alignment = align_right
    c_st_fin.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_subtot_fin, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_subtot_fin, column=c).border = border_cell
    ws_sum.row_dimensions[r_subtot_fin].height = 23

    # Row 15% VAT
    r_vat = 28
    ws_sum.merge_cells(f"B{r_vat}:H{r_vat}")
    ws_sum[f"B{r_vat}"] = "15% VAT"
    ws_sum[f"B{r_vat}"].font = font_data_bold
    ws_sum[f"B{r_vat}"].alignment = align_right
    c_vat_val = ws_sum.cell(row=r_vat, column=9, value="=I27*'Rates & Parameters'!$F$23")
    c_vat_val.font = font_data_bold
    c_vat_val.alignment = align_right
    c_vat_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_vat, column=c).border = border_cell
    ws_sum.row_dimensions[r_vat].height = 22

    # Row Total including VAT
    r_tot_vat = 29
    ws_sum.merge_cells(f"B{r_tot_vat}:H{r_tot_vat}")
    ws_sum[f"B{r_tot_vat}"] = "Total including VAT"
    ws_sum[f"B{r_tot_vat}"].font = font_data_bold
    ws_sum[f"B{r_tot_vat}"].alignment = align_right
    c_tot_vat_val = ws_sum.cell(row=r_tot_vat, column=9, value="=I27+I28")
    c_tot_vat_val.font = font_data_bold
    c_tot_vat_val.alignment = align_right
    c_tot_vat_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_tot_vat, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_tot_vat, column=c).border = border_cell
    ws_sum.row_dimensions[r_tot_vat].height = 23

    # Row Total Project Cost (A) incl VAT
    r_grand_tender = 30
    ws_sum.merge_cells(f"B{r_grand_tender}:H{r_grand_tender}")
    ws_sum[f"B{r_grand_tender}"] = "Total Project Cost (A) incl VAT"
    ws_sum[f"B{r_grand_tender}"].font = font_grandtotal
    ws_sum[f"B{r_grand_tender}"].alignment = align_right
    c_gt_fin = ws_sum.cell(row=r_grand_tender, column=9, value="=I29")
    c_gt_fin.font = font_grandtotal
    c_gt_fin.alignment = align_right
    c_gt_fin.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_grand_tender, column=c).fill = fill_grandtotal
        ws_sum.cell(row=r_grand_tender, column=c).border = border_grandtotal
    ws_sum.row_dimensions[r_grand_tender].height = 27

    # Table 3: Summary per Activity Across All Phases
    r_act_hdr = 33
    ws_sum.merge_cells(f"B{r_act_hdr}:I{r_act_hdr}")
    ws_sum[f"B{r_act_hdr}"] = "Summary per Activity from Each Phase (Cross-Phase Cost Analysis)"
    ws_sum[f"B{r_act_hdr}"].font = font_sec_hdr
    ws_sum[f"B{r_act_hdr}"].fill = fill_sec_hdr
    ws_sum[f"B{r_act_hdr}"].alignment = align_left
    ws_sum.row_dimensions[r_act_hdr].height = 24

    act_headers = ["Code", "Broad Activity Category", "Scope & Technical Work Summary", "UOM", "Total Plant Qty", "Blended Unit Rate (R)", "Total Direct Amount (R)", "% of Direct Cost"]
    for col_idx, h in enumerate(act_headers, start=2):
        cell = ws_sum.cell(row=r_act_hdr+1, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [2, 5, 6, 9] else (align_right if col_idx in [7, 8] else align_left)
        cell.border = border_cell
    ws_sum.row_dimensions[r_act_hdr+1].height = 25

    act_summary_rows = [
        ("ACT_EST", "Site Establishment", "Phase site establishment, decontamination, mobilization & H&S setup", "Lot", 7, "=H35/F35", "='Activity Summary'!F14", "=H35/$H$41"),
        ("ACT_01", "1. Steel & Building Structures", "Dismantling, cutting up & removal of all structural steel, gantries & trusses", "Structures", 35, "=H36/F36", "='Activity Summary'!G14", "=H36/$H$41"),
        ("ACT_02", "2. Concrete Demolition", "Breakup & stockpile reinforced concrete & brickwork above and ±1.5m below NGL", "m³", 13400, "=H37/F37", "='Activity Summary'!H14", "=H37/$H$41"),
        ("ACT_03", "3. Loading and Hauling", "Loading and hauling clean pulverized rubble & concrete to designated dumping", "R/ton", 32160, "='Rates & Parameters'!$F$13", "='Activity Summary'!I14", "=H38/$H$41"),
        ("ACT_04", "4. Shaping and Levelling", "Final shaping, levelling, compaction and cleaning off rehabilitated area", "m²", 93072, "='Rates & Parameters'!$F$14", "='Activity Summary'!J14", "=H39/$H$41"),
        ("ACT_CIV", "Specialist Civil Works", "Phase 03: Empty process dam, dewater, backfill & compact certified material", "m³", 3500, "='Rates & Parameters'!$F$15", "='Activity Summary'!K14", "=H40/$H$41")
    ]

    for idx, act in enumerate(act_summary_rows, start=r_act_hdr+2):
        ws_sum.cell(row=idx, column=2, value=act[0]).alignment = align_center
        ws_sum.cell(row=idx, column=3, value=act[1]).alignment = align_left
        ws_sum.cell(row=idx, column=4, value=act[2]).alignment = align_left_wrap
        ws_sum.cell(row=idx, column=5, value=act[3]).alignment = align_center
        
        c_q = ws_sum.cell(row=idx, column=6, value=act[4])
        c_q.alignment = align_right
        c_q.number_format = FMT_QTY
        
        c_r = ws_sum.cell(row=idx, column=7, value=act[5])
        c_r.alignment = align_right
        c_r.number_format = FMT_CURRENCY
        
        c_tot = ws_sum.cell(row=idx, column=8, value=act[6])
        c_tot.alignment = align_right
        c_tot.font = font_data_bold
        c_tot.number_format = FMT_CURRENCY

        c_pct = ws_sum.cell(row=idx, column=9, value=act[7])
        c_pct.alignment = align_right
        c_pct.font = font_data_bold
        c_pct.number_format = FMT_PERCENT

        fill_row = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)
        for c in range(2, 10):
            ws_sum.cell(row=idx, column=c).border = border_cell
            if fill_row.fill_type:
                ws_sum.cell(row=idx, column=c).fill = fill_row
        ws_sum.row_dimensions[idx].height = 22

    # Grand Total Activity Summary Row
    r_act_tot = 41
    ws_sum.merge_cells(f"B{r_act_tot}:G{r_act_tot}")
    ws_sum[f"B{r_act_tot}"] = "TOTAL DIRECT ACTIVITY COST (RECONCILED TO PHASE DIRECT TOTAL)"
    ws_sum[f"B{r_act_tot}"].font = font_grandtotal
    ws_sum[f"B{r_act_tot}"].alignment = align_right

    c_act_tot_val = ws_sum.cell(row=r_act_tot, column=8, value=f"=SUM(H35:H40)")
    c_act_tot_val.font = font_grandtotal
    c_act_tot_val.alignment = align_right
    c_act_tot_val.number_format = FMT_CURRENCY

    c_act_pct_tot = ws_sum.cell(row=r_act_tot, column=9, value=f"=SUM(I35:I40)")
    c_act_pct_tot.font = font_grandtotal
    c_act_pct_tot.alignment = align_right
    c_act_pct_tot.number_format = FMT_PERCENT

    for c in range(2, 10):
        ws_sum.cell(row=r_act_tot, column=c).fill = fill_grandtotal
        ws_sum.cell(row=r_act_tot, column=c).border = border_grandtotal
    ws_sum.row_dimensions[r_act_tot].height = 25

    # Column widths for Summary Sheet
    ws_sum.column_dimensions["A"].width = 3
    ws_sum.column_dimensions["B"].width = 10
    ws_sum.column_dimensions["C"].width = 28
    ws_sum.column_dimensions["D"].width = 44
    ws_sum.column_dimensions["E"].width = 16
    ws_sum.column_dimensions["F"].width = 14
    ws_sum.column_dimensions["G"].width = 15
    ws_sum.column_dimensions["H"].width = 16
    ws_sum.column_dimensions["I"].width = 20

    # =========================================================================
    # SHEET 4: ACTIVITY SUMMARY (CROSS-PHASE MATRIX)
    # =========================================================================
    ws_act = wb.create_sheet(title="Activity Summary")
    ws_act.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_act.merge_cells("B2:L2")
    ws_act["B2"] = "WEST GOLD PLANT - CROSS-PHASE ACTIVITY BREAKDOWN MATRIX"
    ws_act["B2"].font = font_main_title
    ws_act["B2"].fill = fill_main_title
    ws_act["B2"].alignment = align_title
    ws_act.row_dimensions[2].height = 32

    ws_act.merge_cells("B3:L3")
    ws_act["B3"] = "Direct Cost Distribution per Activity Across All 7 Phases (Steel, Concrete, Load & Haul, Shaping & Levelling)"
    ws_act["B3"].font = font_sub_title
    ws_act["B3"].fill = fill_sec_hdr
    ws_act["B3"].alignment = align_center
    ws_act.row_dimensions[3].height = 20

    act_matrix_headers = [
        "Item", "Phase Description", "Area (m²)", "Duration (Mos)",
        "Site Est. (Lot)", "1. Steel & Struct. (R)", "2. Concrete (R)",
        "3. Load & Haul (R)", "4. Shaping & Level (R)", "Specialist Works (R)", "Phase Total (R)"
    ]
    for col_idx, h in enumerate(act_matrix_headers, start=2):
        cell = ws_act.cell(row=5, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [2, 4, 5] else (align_right if col_idx >= 6 else align_left)
        cell.border = border_cell
    ws_act.row_dimensions[5].height = 26

    matrix_phase_links = [
        (1, "Phase 01: Smelt House", 10574.00, 4, "='Master BOQ Breakdown'!G6", "='Master BOQ Breakdown'!G7", "='Master BOQ Breakdown'!G8", "='Master BOQ Breakdown'!G9", "='Master BOQ Breakdown'!G10", 0, "='Master BOQ Breakdown'!G11"),
        (2, "Phase 02: Thickener Dams", 10890.00, 3, "='Master BOQ Breakdown'!G12", "='Master BOQ Breakdown'!G13", "='Master BOQ Breakdown'!G14", "='Master BOQ Breakdown'!G15", "='Master BOQ Breakdown'!G16", 0, "='Master BOQ Breakdown'!G17"),
        (3, "Phase 03: Mill Section & Dams", 16823.00, 4, "='Master BOQ Breakdown'!G18", "='Master BOQ Breakdown'!G19", "='Master BOQ Breakdown'!G20", "='Master BOQ Breakdown'!G22", "='Master BOQ Breakdown'!G23", "='Master BOQ Breakdown'!G21", "='Master BOQ Breakdown'!G24"),
        (4, "Phase 04: Silos & Foundations", 12763.00, 4, "='Master BOQ Breakdown'!G25", "='Master BOQ Breakdown'!G26", "='Master BOQ Breakdown'!G27", "='Master BOQ Breakdown'!G28", "='Master BOQ Breakdown'!G29", 0, "='Master BOQ Breakdown'!G30"),
        (5, "Phase 05: Workshops Area", 10073.00, 2, "='Master BOQ Breakdown'!G31", "='Master BOQ Breakdown'!G32", "='Master BOQ Breakdown'!G33", "='Master BOQ Breakdown'!G34", "='Master BOQ Breakdown'!G35", 0, "='Master BOQ Breakdown'!G36"),
        (6, "Phase 06: Conveyors & Columns", 13607.00, 2, "='Master BOQ Breakdown'!G37", "='Master BOQ Breakdown'!G38", "='Master BOQ Breakdown'!G39", "='Master BOQ Breakdown'!G40", "='Master BOQ Breakdown'!G41", 0, "='Master BOQ Breakdown'!G42"),
        (7, "Phase 07: Main Gate & Change Hse", 18342.00, 3, "='Master BOQ Breakdown'!G43", "='Master BOQ Breakdown'!G44", "='Master BOQ Breakdown'!G45", "='Master BOQ Breakdown'!G46", "='Master BOQ Breakdown'!G47", 0, "='Master BOQ Breakdown'!G48")
    ]

    for idx, r in enumerate(matrix_phase_links, start=6):
        ws_act.cell(row=idx, column=2, value=r[0]).alignment = align_center
        ws_act.cell(row=idx, column=3, value=r[1]).alignment = align_left
        
        c_a = ws_act.cell(row=idx, column=4, value=r[2])
        c_a.alignment = align_right
        c_a.number_format = FMT_QTY_DEC
        
        c_d = ws_act.cell(row=idx, column=5, value=r[3])
        c_d.alignment = align_center
        c_d.number_format = FMT_QTY
        
        for col_i, fml in enumerate(r[4:], start=6):
            cell = ws_act.cell(row=idx, column=col_i, value=fml)
            cell.alignment = align_right
            cell.number_format = FMT_CURRENCY
            if col_i == 12:
                cell.font = font_data_bold
            else:
                cell.font = font_data

        fill_row = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)
        for c in range(2, 13):
            ws_act.cell(row=idx, column=c).border = border_cell
            if fill_row.fill_type:
                ws_act.cell(row=idx, column=c).fill = fill_row
        ws_act.row_dimensions[idx].height = 22

    # Grand Totals Row for Matrix
    r_mat_tot = 14
    ws_act.merge_cells(f"B{r_mat_tot}:C{r_mat_tot}")
    ws_act[f"B{r_mat_tot}"] = "TOTAL DIRECT PROJECT COST"
    ws_act[f"B{r_mat_tot}"].font = font_grandtotal
    ws_act[f"B{r_mat_tot}"].alignment = align_right

    c_tot_area = ws_act.cell(row=r_mat_tot, column=4, value="=SUM(D6:D12)")
    c_tot_area.alignment = align_right
    c_tot_area.font = font_grandtotal
    c_tot_area.number_format = FMT_QTY_DEC

    c_tot_dur = ws_act.cell(row=r_mat_tot, column=5, value="=SUM(E6:E12)")
    c_tot_dur.alignment = align_center
    c_tot_dur.font = font_grandtotal
    c_tot_dur.number_format = FMT_QTY

    for col_i, col_letter in enumerate(["F", "G", "H", "I", "J", "K", "L"], start=6):
        cell = ws_act.cell(row=r_mat_tot, column=col_i, value=f"=SUM({col_letter}6:{col_letter}12)")
        cell.alignment = align_right
        cell.font = font_grandtotal
        cell.number_format = FMT_CURRENCY

    for c in range(2, 13):
        ws_act.cell(row=r_mat_tot, column=c).fill = fill_grandtotal
        ws_act.cell(row=r_mat_tot, column=c).border = border_grandtotal
    ws_act.row_dimensions[r_mat_tot].height = 25

    # Percentage Contribution Row
    r_mat_pct = 15
    ws_act.merge_cells(f"B{r_mat_pct}:E{r_mat_pct}")
    ws_act[f"B{r_mat_pct}"] = "% CONTRIBUTION TO DIRECT COST"
    ws_act[f"B{r_mat_pct}"].font = font_subtotal
    ws_act[f"B{r_mat_pct}"].alignment = align_right

    for col_i, col_letter in enumerate(["F", "G", "H", "I", "J", "K"], start=6):
        cell = ws_act.cell(row=r_mat_pct, column=col_i, value=f"={col_letter}14/$L$14")
        cell.alignment = align_right
        cell.font = font_subtotal
        cell.number_format = FMT_PERCENT

    c_tot_pct = ws_act.cell(row=r_mat_pct, column=12, value="=SUM(F15:K15)")
    c_tot_pct.alignment = align_right
    c_tot_pct.font = font_subtotal
    c_tot_pct.number_format = FMT_PERCENT

    for c in range(2, 13):
        ws_act.cell(row=r_mat_pct, column=c).fill = fill_subtotal
        ws_act.cell(row=r_mat_pct, column=c).border = border_subtotal
    ws_act.row_dimensions[r_mat_pct].height = 22

    # Column widths for Activity Summary
    ws_act.column_dimensions["A"].width = 3
    ws_act.column_dimensions["B"].width = 8
    ws_act.column_dimensions["C"].width = 28
    ws_act.column_dimensions["D"].width = 14
    ws_act.column_dimensions["E"].width = 15
    ws_act.column_dimensions["F"].width = 16
    ws_act.column_dimensions["G"].width = 18
    ws_act.column_dimensions["H"].width = 18
    ws_act.column_dimensions["I"].width = 18
    ws_act.column_dimensions["J"].width = 18
    ws_act.column_dimensions["K"].width = 18
    ws_act.column_dimensions["L"].width = 20

    # =========================================================================
    # INDIVIDUAL PHASE SHEETS (PHASE 1 TO PHASE 7 WORK PACKAGES)
    # =========================================================================
    phase_tab_configs = [
        (1, "Phase 1 - Smelt House", "Smelt House & Associated Gold Recovery Plant Area", 10574.00, 4,
         "Breakup & stockpile concrete/brickwork above and ±1.5m below NGL; cut up & remove steel structures; load & haul rubble; shaping & levelling.",
         [
             ("Site Establishment", "Lot", 1, "='Rates & Parameters'!$F$7", "Phase mobilization, decontamination showers, security & environmental compliance setup"),
             ("1. Steel & Structures", "R/structure", 3, "='Rates & Parameters'!$F$8", "Cut up and remove all reinforcing material / structural steel of smelt house"),
             ("2. Concrete Demolition", "m³", 1250, "='Rates & Parameters'!$F$11", "Breakup & stockpile concrete & brickwork above and ±1.5m below NGL"),
             ("3. Loading and Hauling", "R/ton", 3000, "='Rates & Parameters'!$F$13", "Loading and hauling of clean building rubble to designated dumping [3,000 t]"),
             ("4. Shaping and Levelling", "m²", 10574, "='Rates & Parameters'!$F$14", "Shaping / levelling, grading, compaction and cleaning off area [10,574 m²]")
         ]),
        (2, "Phase 2 - Thickener Dams", "Thickener Dams, Rake Tanks & Slurry Infrastructure", 10890.00, 3,
         "Breakup & stockpile thickener concrete walls/floors above & ±1.5m below NGL; dismantle rake bridge steel & launders; load & haul; shaping & levelling.",
         [
             ("Site Establishment", "Lot", 1, "='Rates & Parameters'!$F$7", "Phase mobilization, thickener slurry management & safety setup"),
             ("1. Steel & Structures", "R/structure", 3, "='Rates & Parameters'!$F$8", "Dismantle & cut up thickener rake bridge structures, launders & walkways"),
             ("2. Concrete Demolition", "m³", 1600, "='Rates & Parameters'!$F$11", "Breakup & stockpile concrete & brickwork above & ±1.5m below NGL – Thickener Dams"),
             ("3. Loading and Hauling", "R/ton", 3840, "='Rates & Parameters'!$F$13", "Loading and hauling rubble to designated dumping [3,840 t]"),
             ("4. Shaping and Levelling", "m²", 10890, "='Rates & Parameters'!$F$14", "Shaping / levelling and cleaning off area [10,890 m²]")
         ]),
        (3, "Phase 3 - Mill Section", "Mill Section Heavy Industrial & Process Dam Area", 16823.00, 4,
         "Dismantle heavy mill building steel; break heavy concrete mill plinths above & ±1.5m below NGL; empty & backfill process dams; load & haul; shape & level.",
         [
             ("Site Establishment", "Lot", 1, "='Rates & Parameters'!$F$7", "Heavy mobile crane staging, deep excavation safety & mobilization setup"),
             ("1. Steel & Structures", "R/structure", 6, "='Rates & Parameters'!$F$9", "Dismantling of mill section heavy steel structures, gantry frames & reinforcing"),
             ("2. Concrete Demolition", "m³", 2800, "='Rates & Parameters'!$F$11", "Breakup & stockpile concrete & brickwork above and ±1.5m below NGL incl. mill plinths"),
             ("Specialist Civil Works", "m³", 3500, "='Rates & Parameters'!$F$15", "Empty process dams and backfill with certified materials up ±1.5m below NGL"),
             ("3. Loading and Hauling", "R/ton", 6720, "='Rates & Parameters'!$F$13", "Remove / load / haul pulverized rubble to designated stockpile [6,720 t]"),
             ("4. Shaping and Levelling", "m²", 16823, "='Rates & Parameters'!$F$14", "Shaping / levelling and cleaning off area [16,823 m²]")
         ]),
        (4, "Phase 4 - Silos & Fdns", "Concrete Silos & Heavy Silo Foundations Area", 12763.00, 4,
         "Dismantle high-level silo headframes & feed chutes; demolish concrete silos & foundations above & ±1.5m below NGL; load & haul; shape & level.",
         [
             ("Site Establishment", "Lot", 1, "='Rates & Parameters'!$F$7", "High-elevation rigging exclusion zone, dust suppression & mobilization setup"),
             ("1. Steel & Structures", "R/structure", 5, "='Rates & Parameters'!$F$8", "Dismantling of steel structures, conveyor headframes & high-level chutes"),
             ("2. Concrete Demolition", "m³", 3200, "='Rates & Parameters'!$F$11", "Demolition of concrete silos & breakup foundations above & ±1.5m below NGL"),
             ("3. Loading and Hauling", "R/ton", 7680, "='Rates & Parameters'!$F$13", "Remove / load / haul all pulverized material to designated dumping [7,680 t]"),
             ("4. Shaping and Levelling", "m²", 12763, "='Rates & Parameters'!$F$14", "Shaping / levelling and cleaning off area [12,763 m²]")
         ]),
        (5, "Phase 5 - Workshops", "Workshops, Maintenance Sheds & Crane Bays Area", 10073.00, 2,
         "Dismantle workshop portal frames, overhead cranes & trusses; break concrete slabs/footings above & ±1.5m below NGL; load & haul; shape & level.",
         [
             ("Site Establishment", "Lot", 1, "='Rates & Parameters'!$F$7", "Workshop decontamination, hydrocarbon clearance & mobilization setup"),
             ("1. Steel & Structures", "R/structure", 4, "='Rates & Parameters'!$F$10", "Dismantling of workshops steel structures, portal frames & overhead cranes"),
             ("2. Concrete Demolition", "m³", 1300, "='Rates & Parameters'!$F$11", "Breakup & stockpile concrete & brickwork above and ±1.5m below NGL"),
             ("3. Loading and Hauling", "R/ton", 3120, "='Rates & Parameters'!$F$13", "Remove / load / haul all pulverized material to designated dumping [3,120 t]"),
             ("4. Shaping and Levelling", "m²", 10073, "='Rates & Parameters'!$F$14", "Shaping / levelling and cleaning off area [10,073 m²]")
         ]),
        (6, "Phase 6 - Conveyor Struct", "Overland Conveyor Structures & Concrete Trestles Area", 13607.00, 2,
         "Dismantle conveyor steel gantries & stringers; break concrete trestle columns/counterweights above & ±1.5m below NGL; load & haul; shape & level.",
         [
             ("Site Establishment", "Lot", 1, "='Rates & Parameters'!$F$7", "Corridor access, Eskom Western Reef Substation safety interface & setup"),
             ("1. Steel & Structures", "R/structure", 6, "='Rates & Parameters'!$F$8", "Dismantling of conveyor structure steel structures, gantry frames & trestles"),
             ("2. Concrete Demolition", "m³", 1150, "='Rates & Parameters'!$F$11", "Breaking concrete columns, counterweights and plinths above and ±1.5m below NGL"),
             ("3. Loading and Hauling", "R/ton", 2760, "='Rates & Parameters'!$F$13", "Remove / load / haul all pulverized material to designated dumping [2,760 t]"),
             ("4. Shaping and Levelling", "m²", 13607, "='Rates & Parameters'!$F$14", "Shaping / levelling and cleaning off area [13,607 m²]")
         ]),
        (7, "Phase 7 - Main Gate", "Main Gate Workshops, Change Houses & Administrative Area", 18342.00, 3,
         "Breakup all workshops & change houses from main gate; break concrete & brick masonry above & ±1.5m below NGL; load & haul; shape & level.",
         [
             ("Site Establishment", "Lot", 1, "='Rates & Parameters'!$F$7", "Main gate security, public traffic management & mobilization setup"),
             ("1. Steel & Structures", "R/structure", 8, "='Rates & Parameters'!$F$10", "Breakup all workshops and change houses structures from the main gate"),
             ("2. Concrete Demolition", "m³", 2100, "='Rates & Parameters'!$F$11", "Breakup & stockpile concrete and brickwork above and ±1.5m below NGL"),
             ("3. Loading and Hauling", "R/ton", 5040, "='Rates & Parameters'!$F$13", "Remove / load / haul all pulverized material to designated dumping [5,040 t]"),
             ("4. Shaping and Levelling", "m²", 18342, "='Rates & Parameters'!$F$14", "Shaping / levelling and cleaning off area [18,342 m²]")
         ])
    ]

    for p_cfg in phase_tab_configs:
        p_num, p_sheet_name, p_full_title, p_area, p_dur, p_scope_desc, p_line_items = p_cfg
        ws_p = wb.create_sheet(title=p_sheet_name)
        ws_p.views.sheetView[0].showGridLines = True

        ws_p.merge_cells("B2:G2")
        ws_p["B2"] = f"WEST GOLD PLANT - {p_full_title.upper()}"
        ws_p["B2"].font = font_main_title
        ws_p["B2"].fill = fill_main_title
        ws_p["B2"].alignment = align_title
        ws_p.row_dimensions[2].height = 32

        ws_p.merge_cells("B3:G3")
        ws_p["B3"] = f"Area: {p_area:,.2f} m² | Duration: ±{p_dur} Months | Target Clearing: ±1.5m below Natural Ground Level (NGL)"
        ws_p["B3"].font = font_sub_title
        ws_p["B3"].fill = fill_sec_hdr
        ws_p["B3"].alignment = align_center
        ws_p.row_dimensions[3].height = 20

        ws_p.merge_cells("B4:G4")
        ws_p["B4"] = f"SCOPE OF WORK: {p_scope_desc}"
        ws_p["B4"].font = Font(name="Segoe UI", size=9, italic=True, color="1B365D")
        ws_p["B4"].alignment = align_left_wrap
        ws_p.row_dimensions[4].height = 24

        p_tbl_hdrs = ["Item Ref", "Activity & Detailed Description", "UOM", "Quantity", "Unit Rate (R)", "Total (R)"]
        for col_idx, h in enumerate(p_tbl_hdrs, start=2):
            cell = ws_p.cell(row=6, column=col_idx, value=h)
            cell.font = font_tbl_hdr
            cell.fill = fill_main_title
            cell.alignment = align_center if col_idx in [2, 4, 5] else (align_right if col_idx in [6, 7] else align_left)
            cell.border = border_cell
        ws_p.row_dimensions[6].height = 25

        for l_idx, litm in enumerate(p_line_items, start=7):
            cat, uom, qty, rate_fml, desc_long = litm
            ws_p.cell(row=l_idx, column=2, value=f"{p_num}.{l_idx-6}").alignment = align_center
            ws_p.cell(row=l_idx, column=3, value=f"{cat}: {desc_long}").alignment = align_left_wrap
            ws_p.cell(row=l_idx, column=4, value=uom).alignment = align_center
            
            c_q = ws_p.cell(row=l_idx, column=5, value=qty)
            c_q.alignment = align_right
            c_q.number_format = FMT_QTY if isinstance(qty, int) else FMT_QTY_DEC
            
            c_r = ws_p.cell(row=l_idx, column=6, value=rate_fml)
            c_r.alignment = align_right
            c_r.number_format = FMT_CURRENCY
            
            c_t = ws_p.cell(row=l_idx, column=7, value=f"=E{l_idx}*F{l_idx}")
            c_t.alignment = align_right
            c_t.font = font_data_bold
            c_t.number_format = FMT_CURRENCY

            for c in range(2, 8):
                ws_p.cell(row=l_idx, column=c).border = border_cell
                ws_p.cell(row=l_idx, column=c).font = font_data if c != 7 else font_data_bold
            ws_p.row_dimensions[l_idx].height = 24

        r_p_sub = 7 + len(p_line_items)
        ws_p.merge_cells(f"B{r_p_sub}:E{r_p_sub}")
        ws_p[f"B{r_p_sub}"] = f"SUB-TOTAL: {p_sheet_name.upper()} DIRECT DEMOLITION COST"
        ws_p[f"B{r_p_sub}"].font = font_subtotal
        ws_p[f"B{r_p_sub}"].alignment = align_right
        
        ws_p.cell(row=r_p_sub, column=6, value="Phase Subtotal").alignment = align_right
        ws_p.cell(row=r_p_sub, column=6).font = font_subtotal

        c_p_sub_val = ws_p.cell(row=r_p_sub, column=7, value=f"=SUM(G7:G{r_p_sub-1})")
        c_p_sub_val.alignment = align_right
        c_p_sub_val.font = font_subtotal
        c_p_sub_val.number_format = FMT_CURRENCY

        for c in range(2, 8):
            ws_p.cell(row=r_p_sub, column=c).fill = fill_subtotal
            ws_p.cell(row=r_p_sub, column=c).border = border_subtotal
        ws_p.row_dimensions[r_p_sub].height = 25

        r_kpi_start = r_p_sub + 2
        ws_p.merge_cells(f"B{r_kpi_start}:D{r_kpi_start}")
        ws_p[f"B{r_kpi_start}"] = "PHASE PERFORMANCE INDICATORS"
        ws_p[f"B{r_kpi_start}"].font = font_sec_hdr
        ws_p[f"B{r_kpi_start}"].fill = fill_sec_hdr
        ws_p[f"B{r_kpi_start}"].alignment = align_left

        kpis_phase = [
            ("Direct Cost per Footprint Area (R/m²)", f"=G{r_p_sub}/{p_area}", FMT_CURRENCY),
            ("Monthly Direct Expenditure Burn Rate", f"=G{r_p_sub}/{p_dur}", FMT_CURRENCY),
            ("Rubble Generation Rate per Footprint (t/m²)", f"=E{10 if p_num != 3 else 11}/{p_area}", '0.00 "t/m²"')
        ]
        for k_idx, k_item in enumerate(kpis_phase, start=r_kpi_start+1):
            k_lbl, k_fml, k_fmt = k_item
            ws_p.merge_cells(f"B{k_idx}:C{k_idx}")
            ws_p[f"B{k_idx}"] = k_lbl
            ws_p[f"B{k_idx}"].font = font_data
            ws_p[f"B{k_idx}"].border = border_cell
            
            c_k_v = ws_p.cell(row=k_idx, column=4, value=k_fml)
            c_k_v.font = font_data_bold
            c_k_v.alignment = align_right
            c_k_v.number_format = k_fmt
            c_k_v.border = border_cell
            ws_p.row_dimensions[k_idx].height = 20

        ws_p.column_dimensions["A"].width = 3
        ws_p.column_dimensions["B"].width = 12
        ws_p.column_dimensions["C"].width = 48
        ws_p.column_dimensions["D"].width = 14
        ws_p.column_dimensions["E"].width = 16
        ws_p.column_dimensions["F"].width = 18
        ws_p.column_dimensions["G"].width = 22

    # =========================================================================
    # SHEET 5: PROJECT SCHEDULE & CASH FLOW (22 MONTHS S-CURVE)
    # =========================================================================
    ws_sched = wb.create_sheet(title="Schedule & Cashflow")
    ws_sched.views.sheetView[0].showGridLines = True

    ws_sched.merge_cells("B2:Z2")
    ws_sched["B2"] = "WEST GOLD PLANT - 22-MONTH PROJECT TIMELINE & CASH FLOW FORECAST"
    ws_sched["B2"].font = font_main_title
    ws_sched["B2"].fill = fill_main_title
    ws_sched["B2"].alignment = align_title
    ws_sched.row_dimensions[2].height = 32

    ws_sched.merge_cells("B3:Z3")
    ws_sched["B3"] = "Phased Execution Plan, Monthly Billing Distribution, Markups Allocation & Cumulative S-Curve"
    ws_sched["B3"].font = font_sub_title
    ws_sched["B3"].fill = fill_sec_hdr
    ws_sched["B3"].alignment = align_center
    ws_sched.row_dimensions[3].height = 20

    sched_hdrs = ["Item", "Phase / Contract Component", "Total Cost (R)"] + [f"M{m:02d}" for m in range(1, 23)]
    for col_idx, h in enumerate(sched_hdrs, start=2):
        cell = ws_sched.cell(row=5, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center
        cell.border = border_cell
    ws_sched.row_dimensions[5].height = 25

    sched_schedule_data = [
        (1, "Phase 01: Smelt House Area", "='Executive Summary & Financials'!I10", {1: 0.25, 2: 0.25, 3: 0.25, 4: 0.25}),
        (2, "Phase 02: Thickener Dams Area", "='Executive Summary & Financials'!I11", {4: 0.3333, 5: 0.3333, 6: 0.3334}),
        (3, "Phase 03: Mill Section & Dams Area", "='Executive Summary & Financials'!I12", {6: 0.25, 7: 0.25, 8: 0.25, 9: 0.25}),
        (4, "Phase 04: Silos & Foundations Area", "='Executive Summary & Financials'!I13", {9: 0.25, 10: 0.25, 11: 0.25, 12: 0.25}),
        (5, "Phase 05: Workshops Area", "='Executive Summary & Financials'!I14", {13: 0.50, 14: 0.50}),
        (6, "Phase 06: Conveyors & Columns", "='Executive Summary & Financials'!I15", {15: 0.50, 16: 0.50}),
        (7, "Phase 07: Main Gate & Change Hse", "='Executive Summary & Financials'!I16", {17: 0.3333, 18: 0.3333, 19: 0.3334})
    ]

    for idx, s in enumerate(sched_schedule_data, start=6):
        item_i, name_s, tot_fml, dist_map = s
        ws_sched.cell(row=idx, column=2, value=item_i).alignment = align_center
        ws_sched.cell(row=idx, column=3, value=name_s).alignment = align_left
        
        c_tot = ws_sched.cell(row=idx, column=4, value=tot_fml)
        c_tot.alignment = align_right
        c_tot.font = font_data_bold
        c_tot.number_format = FMT_CURRENCY

        for m in range(1, 23):
            col_m = 4 + m
            pct = dist_map.get(m, 0.0)
            cell_m = ws_sched.cell(row=idx, column=col_m)
            if pct > 0:
                cell_m.value = f"=$D{idx}*{pct}"
                cell_m.font = font_data
                cell_m.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            else:
                cell_m.value = 0.0
                cell_m.font = Font(name="Segoe UI", size=9, color="A6A6A6")
            cell_m.alignment = align_right
            cell_m.number_format = FMT_CURRENCY
            cell_m.border = border_cell

        for c in [2, 3, 4]:
            ws_sched.cell(row=idx, column=c).border = border_cell
        ws_sched.row_dimensions[idx].height = 22

    # Monthly Direct Cost Row
    r_mo_dir = 13
    ws_sched.merge_cells(f"B{r_mo_dir}:C{r_mo_dir}")
    ws_sched[f"B{r_mo_dir}"] = "MONTHLY DIRECT PROJECT SPEND"
    ws_sched[f"B{r_mo_dir}"].font = font_subtotal
    ws_sched[f"B{r_mo_dir}"].alignment = align_right

    c_dir_tot = ws_sched.cell(row=r_mo_dir, column=4, value="=SUM(D6:D12)")
    c_dir_tot.alignment = align_right
    c_dir_tot.font = font_subtotal
    c_dir_tot.number_format = FMT_CURRENCY

    for m in range(1, 23):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_dir, column=col_m, value=f"=SUM({col_letter}6:{col_letter}12)")
        cell.alignment = align_right
        cell.font = font_subtotal
        cell.number_format = FMT_CURRENCY

    for c in range(2, 27):
        ws_sched.cell(row=r_mo_dir, column=c).fill = fill_subtotal
        ws_sched.cell(row=r_mo_dir, column=c).border = border_subtotal
    ws_sched.row_dimensions[r_mo_dir].height = 23

    # Monthly Total Billing Row
    r_mo_bill = 14
    ws_sched.merge_cells(f"B{r_mo_bill}:C{r_mo_bill}")
    ws_sched[f"B{r_mo_bill}"] = "TOTAL MONTHLY TENDER BILLING (INCL. P&G, CONT., PROFIT & VAT)"
    ws_sched[f"B{r_mo_bill}"].font = font_grandtotal
    ws_sched[f"B{r_mo_bill}"].alignment = align_right

    c_bill_tot = ws_sched.cell(row=r_mo_bill, column=4, value="='Executive Summary & Financials'!I30")
    c_bill_tot.alignment = align_right
    c_bill_tot.font = font_grandtotal
    c_bill_tot.number_format = FMT_CURRENCY

    for m in range(1, 23):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_bill, column=col_m, value=f"={col_letter}13*($D$14/$D$13)")
        cell.alignment = align_right
        cell.font = font_grandtotal
        cell.number_format = FMT_CURRENCY

    for c in range(2, 27):
        ws_sched.cell(row=r_mo_bill, column=c).fill = fill_grandtotal
        ws_sched.cell(row=r_mo_bill, column=c).border = border_grandtotal
    ws_sched.row_dimensions[r_mo_bill].height = 25

    # Cumulative S-Curve Cash Flow Row
    r_cum = 15
    ws_sched.merge_cells(f"B{r_cum}:C{r_cum}")
    ws_sched[f"B{r_cum}"] = "CUMULATIVE CASH FLOW S-CURVE (INCL. VAT)"
    ws_sched[f"B{r_cum}"].font = font_subtotal
    ws_sched[f"B{r_cum}"].alignment = align_right

    ws_sched.cell(row=r_cum, column=4, value="='Executive Summary & Financials'!I30").number_format = FMT_CURRENCY
    ws_sched.cell(row=r_cum, column=4).font = font_subtotal
    ws_sched.cell(row=r_cum, column=4).alignment = align_right

    for m in range(1, 23):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        if m == 1:
            cell = ws_sched.cell(row=r_cum, column=col_m, value=f"={col_letter}14")
        else:
            prev_letter = get_column_letter(col_m - 1)
            cell = ws_sched.cell(row=r_cum, column=col_m, value=f"={prev_letter}15+{col_letter}14")
        cell.alignment = align_right
        cell.font = font_subtotal
        cell.number_format = FMT_CURRENCY

    for c in range(2, 27):
        ws_sched.cell(row=r_cum, column=c).fill = fill_subtotal
        ws_sched.cell(row=r_cum, column=c).border = border_subtotal
    ws_sched.row_dimensions[r_cum].height = 23

    # Column widths for Schedule
    ws_sched.column_dimensions["A"].width = 3
    ws_sched.column_dimensions["B"].width = 8
    ws_sched.column_dimensions["C"].width = 30
    ws_sched.column_dimensions["D"].width = 20
    for m in range(1, 23):
        col_letter = get_column_letter(4 + m)
        ws_sched.column_dimensions[col_letter].width = 16

    # =========================================================================
    # SHEET 6: PROJECT MAP & SCOPE OF WORK
    # =========================================================================
    ws_map = wb.create_sheet(title="Project Map & Scope")
    ws_map.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_map.merge_cells("B2:H2")
    ws_map["B2"] = "WEST GOLD PLANT - SITE AERIAL MAP & DETAILED SCOPE OF WORK"
    ws_map["B2"].font = font_main_title
    ws_map["B2"].fill = fill_main_title
    ws_map["B2"].alignment = align_title
    ws_map.row_dimensions[2].height = 32

    ws_map.merge_cells("B3:H3")
    ws_map["B3"] = "Total Plant Footprint: 93,072.00 m² | 7 Operational Phases | Depth: ±1.5m below Natural Ground Level"
    ws_map["B3"].font = font_sub_title
    ws_map["B3"].fill = fill_sec_hdr
    ws_map["B3"].alignment = align_center
    ws_map.row_dimensions[3].height = 20

    # Embed Aerial Map Image
    img_path = r"C:/Users/thoso/.gemini/antigravity/brain/2745eacb-735b-4ec2-a4e0-63d91dd6b511/.user_uploaded/media_1793076931680_e35dc9dc.png"
    if os.path.exists(img_path):
        try:
            img = OpenpyxlImage(img_path)
            img.width = 680
            img.height = 355
            ws_map.add_image(img, "B5")
        except Exception as e:
            print(f"Image insert notice: {e}")

    # Scope Table below image
    r_map_scope = 25
    ws_map.merge_cells(f"B{r_map_scope}:H{r_map_scope}")
    ws_map[f"B{r_map_scope}"] = "WEST GOLD PLANT - DETAILED SCOPE OF WORK BY OPERATIONAL PHASE"
    ws_map[f"B{r_map_scope}"].font = font_sec_hdr
    ws_map[f"B{r_map_scope}"].fill = fill_sec_hdr
    ws_map[f"B{r_map_scope}"].alignment = align_left
    ws_map.row_dimensions[r_map_scope].height = 24

    map_scope_hdrs = ["Phase", "Surface Area (m²)", "Duration", "Contractual Demolition & Rehabilitation Scope Specifications"]
    ws_map.cell(row=r_map_scope+1, column=2, value=map_scope_hdrs[0]).font = font_tbl_hdr
    ws_map.cell(row=r_map_scope+1, column=2).fill = fill_main_title
    ws_map.cell(row=r_map_scope+1, column=2).alignment = align_center
    ws_map.cell(row=r_map_scope+1, column=2).border = border_cell

    ws_map.cell(row=r_map_scope+1, column=3, value=map_scope_hdrs[1]).font = font_tbl_hdr
    ws_map.cell(row=r_map_scope+1, column=3).fill = fill_main_title
    ws_map.cell(row=r_map_scope+1, column=3).alignment = align_right
    ws_map.cell(row=r_map_scope+1, column=3).border = border_cell

    ws_map.cell(row=r_map_scope+1, column=4, value=map_scope_hdrs[2]).font = font_tbl_hdr
    ws_map.cell(row=r_map_scope+1, column=4).fill = fill_main_title
    ws_map.cell(row=r_map_scope+1, column=4).alignment = align_center
    ws_map.cell(row=r_map_scope+1, column=4).border = border_cell

    ws_map.merge_cells(f"E{r_map_scope+1}:H{r_map_scope+1}")
    ws_map[f"E{r_map_scope+1}"] = map_scope_hdrs[3]
    ws_map[f"E{r_map_scope+1}"].font = font_tbl_hdr
    ws_map[f"E{r_map_scope+1}"].fill = fill_main_title
    ws_map[f"E{r_map_scope+1}"].alignment = align_left
    for c in range(5, 9):
        ws_map.cell(row=r_map_scope+1, column=c).border = border_cell
    ws_map.row_dimensions[r_map_scope+1].height = 25

    map_scope_rows = [
        ("Phase 01", 10574.00, "+-4 Months", "Demolition and Rehabilitation of West Gold Plant – Phase 01: Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Cut up and remove all reinforcing material / steel structures of smelt house; Remove / Load / Haul all material and stockpile at designated area; Shaping / Levelling and cleaning off area."),
        ("Phase 02", 10890.00, "+-3 Months", "Demolition and Rehabilitation of West Gold Plant – Phase 02: Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level – Thickener Dames; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all material and stockpile at designated area; Shaping / Levelling and cleaning off area."),
        ("Phase 03", 16823.00, "+-4 Months", "Demolition and Rehabilitation of West Gold Plant – Phase 03: Dismantling of mill section steel structures and breaking concrete; Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level including the Dam on-site; Empty dams and backfill with materials up ±1.5 meters below natural ground; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all pulverized material and stockpile at designated area; Shaping / Levelling and cleaning off area."),
        ("Phase 04", 12763.00, "+-4 Months", "Demolition and Rehabilitation of West Gold Plant – Phase 04: Dismantling of steel structures and demolition of concrete silos; Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level – Silo Foundations; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all pulverized material and stockpile at designated area; Shaping / Levelling and cleaning off area."),
        ("Phase 05", 10073.00, "+-2 Months", "Demolition and Rehabilitation of West Gold Plant – Phase 05: Dismantling of workshops steel structures; Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all pulverized material and stockpile at designated area; Shaping / Levelling and cleaning off area."),
        ("Phase 06", 13607.00, "+-2 Months", "Demolition and Rehabilitation of West Gold Plant – Phase 06: Dismantling of conveyor structure steel structures and breaking concrete columns; Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all pulverized material and stockpile at designated area; Shaping / Levelling and cleaning off area."),
        ("Phase 07", 18342.00, "+-3 Months", "Demolition and Rehabilitation of West Gold Plant – Phase 07: Breakup all workshops and change houses structures from the main gate; Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all pulverized material and stockpile at designated area; Shaping / Levelling and cleaning off area.")
    ]

    for idx, ms in enumerate(map_scope_rows, start=r_map_scope+2):
        ws_map.cell(row=idx, column=2, value=ms[0]).alignment = align_center
        ws_map.cell(row=idx, column=2).font = font_data_bold
        
        c_a = ws_map.cell(row=idx, column=3, value=ms[1])
        c_a.alignment = align_right
        c_a.font = font_data
        c_a.number_format = FMT_QTY_DEC
        
        ws_map.cell(row=idx, column=4, value=ms[2]).alignment = align_center
        ws_map.cell(row=idx, column=4).font = font_data
        
        ws_map.merge_cells(f"E{idx}:H{idx}")
        ws_map[f"E{idx}"] = ms[3]
        ws_map[f"E{idx}"].font = font_data
        ws_map[f"E{idx}"].alignment = align_left_wrap

        fill_row = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)
        for c in range(2, 9):
            ws_map.cell(row=idx, column=c).border = border_cell
            if fill_row.fill_type:
                ws_map.cell(row=idx, column=c).fill = fill_row
        ws_map.row_dimensions[idx].height = 40

    r_map_tot = r_map_scope + 2 + len(map_scope_rows)
    ws_map.cell(row=r_map_tot, column=2, value="TOTAL").alignment = align_center
    ws_map.cell(row=r_map_tot, column=2).font = font_grandtotal

    c_tot_map_area = ws_map.cell(row=r_map_tot, column=3, value=f"=SUM(C{r_map_scope+2}:C{r_map_tot-1})")
    c_tot_map_area.alignment = align_right
    c_tot_map_area.font = font_grandtotal
    c_tot_map_area.number_format = FMT_QTY_DEC

    ws_map.cell(row=r_map_tot, column=4, value="+-22 Months").alignment = align_center
    ws_map.cell(row=r_map_tot, column=4).font = font_grandtotal

    ws_map.merge_cells(f"E{r_map_tot}:H{r_map_tot}")
    ws_map[f"E{r_map_tot}"] = "Complete West Gold Plant Footprint Rehabilitation (Target Depth: ±1.5m below NGL)"
    ws_map[f"E{r_map_tot}"].font = font_grandtotal
    ws_map[f"E{r_map_tot}"].alignment = align_left

    for c in range(2, 9):
        ws_map.cell(row=r_map_tot, column=c).fill = fill_grandtotal
        ws_map.cell(row=r_map_tot, column=c).border = border_grandtotal
    ws_map.row_dimensions[r_map_tot].height = 26

    # Column widths for Map & Scope
    ws_map.column_dimensions["A"].width = 3
    ws_map.column_dimensions["B"].width = 14
    ws_map.column_dimensions["C"].width = 18
    ws_map.column_dimensions["D"].width = 16
    ws_map.column_dimensions["E"].width = 24
    ws_map.column_dimensions["F"].width = 24
    ws_map.column_dimensions["G"].width = 24
    ws_map.column_dimensions["H"].width = 24

    wb.save(output_filename)
    print(f"Workbook successfully saved to: {output_filename}")

if __name__ == "__main__":
    build_complete_west_gold_model()
