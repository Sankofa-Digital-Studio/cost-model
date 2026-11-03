import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage

def build_complete_capm_tau_mine_model(output_filename="CAPM_Tau_Mine_10_Demolition_Rehabilitation_Model.xlsx"):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

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
    ws_rates["B2"] = "CAPM TAU MINE #10 - MASTER UNIT RATES & COMMERCIAL PARAMETERS"
    ws_rates["B2"].font = font_main_title
    ws_rates["B2"].fill = fill_main_title
    ws_rates["B2"].alignment = align_title
    ws_rates.row_dimensions[2].height = 34

    ws_rates.merge_cells("B3:G3")
    ws_rates["B3"] = "CENTRAL RATE REPOSITORY: Edit any rate in the yellow cells below. All 9 Phase sheets, Master BOQ, Summaries and Tender Price update automatically."
    ws_rates["B3"].font = font_sub_title
    ws_rates["B3"].fill = fill_sec_hdr
    ws_rates["B3"].alignment = align_center
    ws_rates.row_dimensions[3].height = 20

    # Section 1: Core Activity Unit Rates
    ws_rates.merge_cells("B5:G5")
    ws_rates["B5"] = "1. CORE DEMOLITION, EARTHWORKS & REHABILITATION MASTER UNIT RATES"
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

    # Core rates data: Row 7 to 18
    # Rates mappings:
    # Row 7: R_EST (Site Est) -> F7
    # Row 8: R_STL_STD (Standard Steel) -> F8
    # Row 9: R_STL_HVY (Heavy Steel Headgears/Silos/Conveyors) -> F9
    # Row 10: R_STL_WKP (Workshops/Hostel Framing) -> F10
    # Row 11: R_CON_M3 (Concrete Breakup m3) -> F11
    # Row 12: R_CON_TON (Concrete Breakup ton) -> F12
    # Row 13: R_HAUL_TON (Loading & Hauling Rubble R/ton) -> F13
    # Row 14: R_SHAPE_M2 (Shaping & Levelling R/m2) -> F14
    # Row 15: R_SILO_DEM (High Silo Demolition R/m3) -> F15
    # Row 16: R_WRD_PROC (Waste Rock Dump Processing R/ton) -> F16
    # Row 17: R_SHAFT_FILL (Shaft Fill Haulage & Stockpiling R/ton) -> F17
    # Row 18: R_SHAFT_CAP (Reinforced Concrete Shaft Capping R/m3) -> F18
    rates_data = [
        ("R_EST", "Site Establishment", "Phase Site Establishment, Mobilization, Health & Safety, Decontamination & Staging Setup", "Lot", 35000.00, "Matches contractual screenshot rate (R 35,000.00 / Lot)"),
        ("R_STL_STD", "1. Steel & Structures", "Dismantling, cutting up & removal of standard structural steel / office / salvage framing", "R/structure", 21000.00, "Matches contractual screenshot rate (R 21,000.00 / structure)"),
        ("R_STL_HVY", "1. Steel & Structures", "Heavy industrial structural steel dismantling (Headgears, winder houses, conveyor gantries)", "R/structure", 28500.00, "Heavy plant structure allowance (high-elevation rigging & flame cutting)"),
        ("R_STL_WKP", "1. Steel & Structures", "Light workshop portal frames, hostel roofing trusses & electrical substation frames", "R/structure", 19500.00, "Matches contractual screenshot rate (R 19,500.00 / structure)"),
        ("R_CON_M3", "2. Concrete Demolition", "Breakup & stockpile reinforced concrete & brickwork above and ±1.5m below NGL", "m³", 240.00, "Heavy excavator hydraulic hammering & mechanical pulverizing per m³"),
        ("R_CON_TON", "2. Concrete Demolition", "Concrete Breakup & Stockpile on tonnage basis (equivalent @ 2.40 t/m³ density)", "R/ton", 100.00, "Tonnage equivalent rate for concrete processing"),
        ("R_HAUL_TON", "3. Loading & Hauling", "Loading and hauling of clean building rubble / pulverized material to designated dumping", "R/ton", 120.00, "Matches contractual screenshot rate exactly (R 120.00 / ton)"),
        ("R_SHAPE_M2", "4. Shaping & Levelling", "Final shaping, levelling, grading, compaction and cleaning off rehabilitated area", "R/m²", 18.50, "Dozer / grader shaping and surface clearance per m² footprint"),
        ("R_SILO_DEM", "Specialist Demolition", "High-elevation concrete silo demolition, specialized shearing & foundation breakup", "m³", 260.00, "Phase 6 specific scope: 2x tall concrete silos & deep raft foundations"),
        ("R_WRD_PROC", "Waste Rock Processing", "Waste rock dump mechanical processing, crushing, screening & sizing for backfill", "R/ton", 45.00, "Phase 8 specific scope: 1,538,058 t processed for shaft backfill material"),
        ("R_SHAFT_FILL", "Shaft Filling Haulage", "Transport all material next to shafts and stockpile for purpose of shaft filling", "R/ton", 55.00, "Contract Scope: Haulage & stockpiling next to Main & RV Shafts (1,538,058 t)"),
        ("R_SHAFT_CAP", "Shaft Concrete Capping", "Capping of deep shafts with heavy reinforced structural concrete (formwork, rebar, 40MPa)", "m³", 4200.00, "Contract Scope: Main Shaft (288 m³) & RV Shaft (392 m³) reinforced capping")
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

    # Section 2: Commercial Markups & Financial Model Parameters (Row 20 to 25)
    # Row 22: P_PG -> F22
    # Row 23: P_CONT -> F23
    # Row 24: P_PROFIT -> F24
    # Row 25: P_VAT -> F25
    ws_rates.merge_cells("B20:G20")
    ws_rates["B20"] = "2. CONTRACT COMMERCIAL MARKUPS & FINANCIAL MODEL PARAMETERS"
    ws_rates["B20"].font = font_sec_hdr
    ws_rates["B20"].fill = fill_sec_hdr
    ws_rates["B20"].alignment = align_left
    ws_rates.row_dimensions[20].height = 24

    comm_headers = ["Parameter Code", "Commercial Parameter Name", "Contractual Description / Basis", "UOM", "Master Parameter Value", "Notes & Reference Grounding"]
    for col_idx, h in enumerate(comm_headers, start=2):
        cell = ws_rates.cell(row=21, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [2, 5, 6] else align_left
        cell.border = border_cell
    ws_rates.row_dimensions[21].height = 25

    comm_params = [
        ("P_PG", "Preliminaries and General (P&G)", "Site management, environmental compliance, health & safety, security, insurances", "%", 0.10, "10.0% P&G (Matches contractual screenshot row: 'Preliminaries and General (10%)')"),
        ("P_CONT", "Contingencies Allowance", "Unforeseen physical conditions, sub-surface shaft variations & weather delays", "%", 0.10, "10.0% Contingencies (Matches contractual screenshot row: 'Contingencies (10%)')"),
        ("P_PROFIT", "Contractor Profit Margin", "Target contractor net operating profit margin applied to Project Costs A", "%", 0.05, "5.0% Profit (Matches contractual screenshot highlighted box: 'Profit Percentage 5%')"),
        ("P_VAT", "Value Added Tax (VAT)", "South African Revenue Service (SARS) standard statutory VAT rate", "%", 0.15, "15.0% VAT (Matches contractual screenshot row: '15% VAT')")
    ]

    for idx, p in enumerate(comm_params, start=22):
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

    # Section 3: Technical Assumptions & Conversion Factors (Row 27 to 35)
    # Row 29: F_DENSITY_CON -> F29
    # Row 30: F_DENSITY_FILL -> F30
    # Row 31: F_BULK -> F31
    # Row 32: F_DEPTH -> F32
    # Row 33: P_TOTAL_AREA -> F33 (1,463,268 m²)
    # Row 34: P_DURATION -> F34 (24.00 Months)
    # Row 35: P_SHAFT_TONS -> F35 (1,538,058 tons)
    ws_rates.merge_cells("B27:G27")
    ws_rates["B27"] = "3. TECHNICAL ENGINEERING ASSUMPTIONS & MATERIAL CONVERSION FACTORS"
    ws_rates["B27"].font = font_sec_hdr
    ws_rates["B27"].fill = fill_sec_hdr
    ws_rates["B27"].alignment = align_left
    ws_rates.row_dimensions[27].height = 24

    tech_headers = ["Factor Code", "Technical Parameter", "Engineering Scope & Physical Definition", "Unit", "Factor Value", "Conversion Application"]
    for col_idx, h in enumerate(tech_headers, start=2):
        cell = ws_rates.cell(row=28, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [2, 5, 6] else align_left
        cell.border = border_cell
    ws_rates.row_dimensions[28].height = 25

    tech_params = [
        ("F_DENSITY_CON", "Reinforced Concrete Density", "Bulk density of reinforced industrial concrete foundations & slabs", "ton/m³", 2.40, "Converts concrete demolition volume (m³) to haulage tonnage (ton)"),
        ("F_DENSITY_FILL", "Shaft Filling Material Density", "Bulk density factor for shaft filling material as specified in contract scope", "ton/m³", 1.70, "Specified in contract docx: Volume x 1.7 = Total Tonnage"),
        ("F_BULK", "Bulking Swell Factor", "Mechanical fragmentation bulking ratio for loose loading and haulage", "ratio", 1.30, "Loose volume check for tipper truck payload calculation"),
        ("F_DEPTH", "Below-Ground Demolition Depth", "Standard rehabilitation clearance depth below natural ground level (NGL)", "meters", 1.50, "Contract requirement: ±1.5m below NGL across all 8 surface phases"),
        ("P_TOTAL_AREA", "Total Mine Surface Footprint", "Sum of Phase 1 through Phase 8 surveyed surface rehabilitation areas", "m²", 1463268.00, "Scope docx exact specification: 'Total area = 146 3268 m²'"),
        ("P_DURATION", "Contract Project Duration", "Integrated project duration spanning 8 surface phases and shaft filling", "Months", 24.00, "Long Term Contract phased timeline with parallel execution"),
        ("P_SHAFT_TONS", "Total Shaft Filling Tonnage", "Combined backfill tonnage for Main Shaft (946,961.2 t) & RV Shaft (591,096.8 t)", "Tonnes", 1538058.00, "Scope docx exact specification: 946,961.2 + 591,096.8 = 1,538,058.0 tons")
    ]

    for idx, t in enumerate(tech_params, start=29):
        ws_rates.cell(row=idx, column=2, value=t[0]).alignment = align_center
        ws_rates.cell(row=idx, column=3, value=t[1]).alignment = align_left
        ws_rates.cell(row=idx, column=4, value=t[2]).alignment = align_left_wrap
        ws_rates.cell(row=idx, column=5, value=t[3]).alignment = align_center
        
        val_cell = ws_rates.cell(row=idx, column=6, value=t[4])
        val_cell.alignment = align_right
        val_cell.font = font_input
        val_cell.fill = fill_input
        val_cell.border = border_input
        val_cell.number_format = FMT_QTY_DEC if t[3] in ["ton/m³", "ratio", "meters", "m²"] else FMT_QTY
        
        ws_rates.cell(row=idx, column=7, value=t[5]).alignment = align_left_wrap
        for c in [2, 3, 4, 5, 7]:
            ws_rates.cell(row=idx, column=c).font = font_data
            ws_rates.cell(row=idx, column=c).border = border_cell
        ws_rates.row_dimensions[idx].height = 22

    # Column widths for Rates & Parameters
    ws_rates.column_dimensions["A"].width = 3
    ws_rates.column_dimensions["B"].width = 16
    ws_rates.column_dimensions["C"].width = 28
    ws_rates.column_dimensions["D"].width = 56
    ws_rates.column_dimensions["E"].width = 14
    ws_rates.column_dimensions["F"].width = 22
    ws_rates.column_dimensions["G"].width = 48

    # =========================================================================
    # SHEET 2: MASTER BOQ BREAKDOWN (CONSOLIDATED SCHEDULE)
    # =========================================================================
    ws_boq = wb.create_sheet(title="Master BOQ Breakdown")
    ws_boq.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_boq.merge_cells("A2:G2")
    ws_boq["A2"] = "CAPM TAU MINE #10 - CONSOLIDATED BILL OF QUANTITIES"
    ws_boq["A2"].font = font_main_title
    ws_boq["A2"].fill = fill_main_title
    ws_boq["A2"].alignment = align_title
    ws_boq.row_dimensions[2].height = 32

    ws_boq.merge_cells("A3:G3")
    ws_boq["A3"] = "Demolition, Concrete Breakup, Load & Haul, Shaping & Levelling, Waste Rock Processing & Shaft Filling (Phases 01 to 08 + Shafts)"
    ws_boq["A3"].font = font_sub_title
    ws_boq["A3"].fill = fill_sec_hdr
    ws_boq["A3"].alignment = align_center
    ws_boq.row_dimensions[3].height = 20

    boq_headers = ["Item", "Mine Phase / Work Package", "Description", "UOM", "Qty", "Unit Rate (R)", "Total (R)"]
    for col_idx, h in enumerate(boq_headers, start=1):
        cell = ws_boq.cell(row=5, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [1, 4, 5] else (align_right if col_idx in [6, 7] else align_left)
        cell.border = border_cell
    ws_boq.row_dimensions[5].height = 26

    # Master Phases Specifications:
    # 8 Surface Phases + 1 Shaft Filling & Capping Work Package
    phases_master = [
        (
            1, "Phase 01: Salvage Yard & Slabs", 178860.00, 3,
            [
                ("Site Establishment, mobilization, health & safety and salvage yard staging setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Cut up and remove all reinforcing material / steel structures & clear salvage yard", "R/structure", 6, "='Rates & Parameters'!$F$8"),
                ("2. Break up all concrete slabs and breakup & stockpile concrete and brickwork above & ±1.5m below NGL", "m³", 4500, "='Rates & Parameters'!$F$11"),
                ("3. Remove / Load / Haul all material and rubble to designated stockpile area (10,800 t)", "R/ton", 10800, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (178,860 m²)", "m²", 178860, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            2, "Phase 02: Admin & Cooling Towers", 102294.00, 4,
            [
                ("Site Establishment, office utilities isolation & decontamination setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Dismantling of workshops steel structures, cooling towers steel frames & cut reinforcing", "R/structure", 8, "='Rates & Parameters'!$F$8"),
                ("2. Breakup all main office buildings, concrete slabs & breaking water reservoir for cooling towers above & ±1.5m below NGL", "m³", 6200, "='Rates & Parameters'!$F$11"),
                ("3. Remove / Load / Haul all material and rubble to designated stockpile area (14,880 t)", "R/ton", 14880, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (102,294 m²)", "m²", 102294, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            3, "Phase 03: Kitchen & Offices Area", 141802.00, 4,
            [
                ("Site Establishment, kitchen isolation, hazardous materials audit & safety setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Cut up and remove all reinforcing material / structural steel frames & roof trusses", "R/structure", 5, "='Rates & Parameters'!$F$8"),
                ("2. Breakup main kitchen and surrounding office buildings concrete & brickwork above & ±1.5m below NGL", "m³", 5800, "='Rates & Parameters'!$F$11"),
                ("3. Remove / Load / Haul all material and building rubble to designated stockpile (13,920 t)", "R/ton", 13920, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (141,802 m²)", "m²", 141802, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            4, "Phase 04: Hostel Complex Area", 12763.00, 4,
            [
                ("Site Establishment, residential block perimeter hoarding & safety setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Cut up and remove all reinforcing material, roof sheeting, trusses & steel structures", "R/structure", 6, "='Rates & Parameters'!$F$10"),
                ("2. Breakup hostel area accommodation blocks, masonry walls and concrete foundations above & ±1.5m below NGL", "m³", 3500, "='Rates & Parameters'!$F$11"),
                ("3. Remove / Load / Haul all material and building rubble to designated stockpile (8,400 t)", "R/ton", 8400, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (12,763 m²)", "m²", 12763, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            5, "Phase 05: Workshops & Substation", 119515.00, 3,
            [
                ("Site Establishment, MV/LV electrical isolation, lockout-tagout & environmental setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Breakup & dismantling of workshops steel structures, gantry frames & cleanup salvage yard", "R/structure", 7, "='Rates & Parameters'!$F$10"),
                ("2. Breakup electrical station yard, heavy plinths, water reservoirs & concrete slabs above & ±1.5m below NGL", "m³", 5200, "='Rates & Parameters'!$F$11"),
                ("3. Remove / Load / Haul all material and pulverized rubble to designated stockpile (12,480 t)", "R/ton", 12480, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (119,515 m²)", "m²", 119515, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            6, "Phase 06: Shafts, Silos & Headgears", 91628.00, 6,
            [
                ("Site Establishment, heavy rigging crane mobilization, shaft collar safety & blast staging setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Demolition & dismantling of two headgear structures (Main Shaft & RV Shaft)", "R/structure", 2, "='Rates & Parameters'!$F$9"),
                ("   Dismantling of conveyor structure steel structures, gantry spans & breaking concrete columns", "R/structure", 4, "='Rates & Parameters'!$F$9"),
                ("   Dismantling steel structures for winder houses, hoist chambers & demolition of vent shaft structure", "R/structure", 4, "='Rates & Parameters'!$F$9"),
                ("2. Breakup two silos' structures & heavy foundation plinths", "m³", 4200, "='Rates & Parameters'!$F$15"),
                ("   Breakup refrigeration plant, winder concrete structures & all concrete slab around shaft area above & ±1.5m below NGL", "m³", 8600, "='Rates & Parameters'!$F$11"),
                ("3. Remove / Load / Haul all pulverized material and heavy debris to designated stockpile (30,720 t)", "R/ton", 30720, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off shaft surface bank area (91,628 m²)", "m²", 91628, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            7, "Phase 07: Dams & Water Reservoirs", 264040.00, 2,
            [
                ("Site Establishment, dewatering pumping mobilization & environmental containment setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Cut up and remove all reinforcing material, steel pipe gantries & pump station structures", "R/structure", 4, "='Rates & Parameters'!$F$8"),
                ("2. Breakup dams and water reservoir concrete walls, sumps and slabs above & ±1.5m below NGL", "m³", 7500, "='Rates & Parameters'!$F$11"),
                ("3. Remove / Load / Haul all material and pulverized rubble to designated stockpile (18,000 t)", "R/ton", 18000, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off area (264,040 m²)", "m²", 264040, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            8, "Phase 08: Waste Rock Dump Process", 552366.00, 12,
            [
                ("Site Establishment, heavy mobile crushing / screening plant mobilization & environmental dust suppression setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Cut up and remove all reinforcing material & steel conveyor discharge structures", "R/structure", 2, "='Rates & Parameters'!$F$8"),
                ("2. Breakup & stockpile concrete footings, transfer plinths and masonry above & ±1.5m below NGL", "m³", 3200, "='Rates & Parameters'!$F$11"),
                ("   Waste rock dump must be processed: crushing, screening & sizing material for shaft backfill (1,538,058 t)", "ton", 1538058, "='Rates & Parameters'!$F$16"),
                ("3. Remove / Load / Haul all concrete rubble to designated stockpile (7,680 t)", "R/ton", 7680, "='Rates & Parameters'!$F$13"),
                ("4. Shaping / Levelling and cleaning off massive waste rock dump footprint area (552,366 m²)", "m²", 552366, "='Rates & Parameters'!$F$14")
            ]
        ),
        (
            9, "Specialist: Shaft Filling & Capping", 0.0, 6,
            [
                ("Site Establishment, deep shaft collar safety barriers, telemetry, heavy winches & environmental setup", "Lot", 1, "='Rates & Parameters'!$F$7"),
                ("1. Transport all material next to Main Shaft and stockpile for shaft filling (10m dia x 1774m deep = 557,036 m³ x 1.7 = 946,961.2 t)", "ton", 946961.2, "='Rates & Parameters'!$F$17"),
                ("2. Capping of the Main Shaft with reinforced concrete (12m x 12m x 2m = 288 m³)", "m³", 288, "='Rates & Parameters'!$F$18"),
                ("3. Transport all material next to RV Vent Shaft and stockpile for shaft filling (14m x 14m x 1774m x 1.7 = 591,096.8 t)", "ton", 591096.8, "='Rates & Parameters'!$F$17"),
                ("4. Capping of the RV Shaft with reinforced steel concrete (14m x 14m x 2m = 392 m³)", "m³", 392, "='Rates & Parameters'!$F$18")
            ]
        )
    ]

    curr_row = 6
    phase_subtotal_rows = {}
    phase_start_end_rows = {}
    boq_item_rows = {}  # Store specific row numbers for Activity Summary cross links

    for p in phases_master:
        item_num, phase_title, area_m2, duration_mo, items = p
        start_row = curr_row
        num_items = len(items)
        end_phase_row = start_row + num_items - 1
        phase_start_end_rows[item_num] = (start_row, end_phase_row)
        boq_item_rows[item_num] = list(range(start_row, end_phase_row + 1))

        # Merge Item Number column
        ws_boq.merge_cells(start_row=start_row, start_column=1, end_row=end_phase_row, end_column=1)
        ws_boq.cell(row=start_row, column=1, value=item_num).alignment = align_center
        ws_boq.cell(row=start_row, column=1).font = font_data_bold

        # Merge Phase Title column
        ws_boq.merge_cells(start_row=start_row, start_column=2, end_row=end_phase_row, end_column=2)
        if area_m2 > 0:
            title_text = f"{phase_title}\n(Area: {area_m2:,.2f} m² | ±{duration_mo} Mos)"
        else:
            title_text = f"{phase_title}\n(1.54M Tons Backfill | 680 m³ Capping | ±{duration_mo} Mos)"
        cell_phase_title = ws_boq.cell(row=start_row, column=2, value=title_text)
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

    # Grand Total Direct Row for Master BOQ
    r_boq_grand = curr_row
    ws_boq.merge_cells(f"A{r_boq_grand}:F{r_boq_grand}")
    ws_boq.cell(row=r_boq_grand, column=1, value="TOTAL DIRECT PROJECT COST (EXCLUDING P&G, CONTINGENCIES & VAT)").alignment = align_right
    ws_boq.cell(row=r_boq_grand, column=1).font = font_grandtotal

    # Sum of all phase subtotal cells
    subtot_cells_str = ",".join([f"G{phase_subtotal_rows[p[0]]}" for p in phases_master])
    c_boq_gt = ws_boq.cell(row=r_boq_grand, column=7, value=f"=SUM({subtot_cells_str})")
    c_boq_gt.alignment = align_right
    c_boq_gt.font = font_grandtotal
    c_boq_gt.number_format = FMT_CURRENCY

    for c in range(1, 8):
        ws_boq.cell(row=r_boq_grand, column=c).fill = fill_grandtotal
        ws_boq.cell(row=r_boq_grand, column=c).border = border_grandtotal
    ws_boq.row_dimensions[r_boq_grand].height = 26

    # Column widths for Master BOQ
    ws_boq.column_dimensions["A"].width = 8
    ws_boq.column_dimensions["B"].width = 30
    ws_boq.column_dimensions["C"].width = 62
    ws_boq.column_dimensions["D"].width = 14
    ws_boq.column_dimensions["E"].width = 16
    ws_boq.column_dimensions["F"].width = 20
    ws_boq.column_dimensions["G"].width = 24

    # =========================================================================
    # SHEET 3: EXECUTIVE SUMMARY & FINANCIALS (WATERFALL & PROFIT MODEL)
    # =========================================================================
    ws_sum = wb.create_sheet(title="Executive Summary & Financials")
    ws_sum.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_sum.merge_cells("B2:I2")
    ws_sum["B2"] = "CAPM TAU MINE #10 - DEMOLITION, SHAFT FILLING & REHABILITATION"
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
    # 4 KPI cards: TOTAL MINE FOOTPRINT, EST. DURATION, TOTAL SHAFT FILLING, CONCRETE DEMOLISHED
    for title, formula, unit, merge_val, merge_lbl in [
        ("TOTAL MINE FOOTPRINT", "='Rates & Parameters'!$F$33", " m²", "B5:C5", "B6:C6"),
        ("EST. DURATION", "='Rates & Parameters'!$F$34", " Months", "D5:E5", "D6:E6"),
        ("TOTAL SHAFT BACKFILL", "='Rates & Parameters'!$F$35", " Tonnes", "F5:G5", "F6:G6"),
        ("TOTAL RUBBLE HAULED", 116880, " Tonnes", "H5:I5", "H6:I6")
    ]:
        ws_sum.merge_cells(merge_val)
        top_left = merge_val.split(":")[0]
        cell_v = ws_sum[top_left]
        cell_v.value = formula
        cell_v.font = font_kpi_num
        cell_v.alignment = align_center
        cell_v.fill = fill_kpi
        if isinstance(formula, str) and "Rates" in formula:
            if "33" in formula:
                cell_v.number_format = '#,##0.00" m²"'
            elif "34" in formula:
                cell_v.number_format = '#,##0" Months"'
            else:
                cell_v.number_format = '#,##0" Tonnes"'
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

    # Table 1: Summary of Demolition, Load and Hauling of Rubble & Shaft Filling
    row_t1_hdr = 8
    ws_sum.merge_cells(f"B{row_t1_hdr}:I{row_t1_hdr}")
    ws_sum[f"B{row_t1_hdr}"] = "Summary of Demolition, Load and Hauling of Rubble & Shaft Filling Packages"
    ws_sum[f"B{row_t1_hdr}"].font = font_sec_hdr
    ws_sum[f"B{row_t1_hdr}"].fill = fill_sec_hdr
    ws_sum[f"B{row_t1_hdr}"].alignment = align_center
    ws_sum.row_dimensions[row_t1_hdr].height = 24

    sum_col_headers = ["Item", "Mine Phase / Plant Section", "Scope Summary", "Footprint Area (m²)", "Duration (Months)", "Structures Qty", "Rubble / Fill (Tonnes)", "Amount (R)"]
    for col_idx, h in enumerate(sum_col_headers, start=2):
        cell = ws_sum.cell(row=row_t1_hdr+1, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [2, 5, 6, 7, 8] else (align_right if col_idx == 9 else align_left)
        cell.border = border_cell
    ws_sum.row_dimensions[row_t1_hdr+1].height = 25

    sum_phase_rows = [
        (1, "Phase 01: Salvage Yard & Slabs Area", "Break up slabs, clear salvage yard, concrete breakup ±1.5m below NGL, load & haul rubble, shape & level", 178860.00, 3, 6, 10800, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[1]}"),
        (2, "Phase 02: Admin & Cooling Towers Area", "Main office buildings, workshop steel, cooling towers & reservoir breakup ±1.5m below NGL, load & haul", 102294.00, 4, 8, 14880, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[2]}"),
        (3, "Phase 03: Kitchen & Offices Area", "Breakup main kitchen, surrounding offices, concrete breakup ±1.5m below NGL, load & haul rubble", 141802.00, 4, 5, 13920, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[3]}"),
        (4, "Phase 04: Hostel Complex Area", "Breakup hostel area accommodation blocks & ablutions ±1.5m below NGL, load & haul rubble, shape & level", 12763.00, 4, 6, 8400, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[4]}"),
        (5, "Phase 05: Workshops & Substation Area", "Dismantle workshops steel, water reservoirs, electrical substation yard breakup ±1.5m below NGL, load & haul", 119515.00, 3, 7, 12480, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[5]}"),
        (6, "Phase 06: Shafts, Silos & Headgears", "Demolish 2 headgears, vent structure, winder houses, 2 concrete silos, breakup slab around shaft, load & haul", 91628.00, 6, 10, 30720, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[6]}"),
        (7, "Phase 07: Dams & Water Reservoirs", "Breakup dams, retaining walls & water reservoir ±1.5m below NGL, cut reinforcing, load & haul rubble", 264040.00, 2, 4, 18000, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[7]}"),
        (8, "Phase 08: Waste Rock Dump Processing", "Waste rock dump processing & crushing (1.54M t), concrete footings breakup ±1.5m below NGL, shape 552,366 m²", 552366.00, 12, 2, 7680, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[8]}"),
        (9, "Specialist: Shaft Filling & Capping", "Haul & stockpile 946,961.2 t to Main Shaft, 591,096.8 t to RV Shaft; 680 m³ reinforced concrete capping", 0.0, 6, 0, 1538058, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[9]}")
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

    # Financial Waterfall Rows (Modeled EXACTLY on West Gold & contract screenshots)
    # Row 19: Sub-Total (Excluding Preliminaries and General) = SUM(I10:I18)
    r_sub_excl_pg = 19
    ws_sum.merge_cells(f"B{r_sub_excl_pg}:H{r_sub_excl_pg}")
    ws_sum.cell(row=r_sub_excl_pg, column=2, value="Sub-Total (Excluding Preliminaries and General)").alignment = align_right
    ws_sum.cell(row=r_sub_excl_pg, column=2).font = font_subtotal

    c_sub_val = ws_sum.cell(row=r_sub_excl_pg, column=9, value="=SUM(I10:I18)")
    c_sub_val.alignment = align_right
    c_sub_val.font = font_subtotal
    c_sub_val.number_format = FMT_CURRENCY

    for c in range(2, 10):
        ws_sum.cell(row=r_sub_excl_pg, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_sub_excl_pg, column=c).border = border_subtotal
    ws_sum.row_dimensions[r_sub_excl_pg].height = 24

    # Row 20: Preliminaries and General (10%)
    r_pg = 20
    ws_sum.merge_cells(f"B{r_pg}:H{r_pg}")
    ws_sum.cell(row=r_pg, column=2, value="Preliminaries and General (10%)").alignment = align_right
    ws_sum.cell(row=r_pg, column=2).font = font_data_bold
    c_pg_val = ws_sum.cell(row=r_pg, column=9, value="=I19*'Rates & Parameters'!$F$22")
    c_pg_val.alignment = align_right
    c_pg_val.font = font_data_bold
    c_pg_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_pg, column=c).border = border_cell
    ws_sum.row_dimensions[r_pg].height = 22

    # Row 21: Sub-Total (Excluding Contingencies)
    r_sub_excl_cont = 21
    ws_sum.merge_cells(f"B{r_sub_excl_cont}:H{r_sub_excl_cont}")
    ws_sum.cell(row=r_sub_excl_cont, column=2, value="Sub-Total (Excluding Contingencies)").alignment = align_right
    ws_sum.cell(row=r_sub_excl_cont, column=2).font = font_subtotal
    c_sub_excl_cont_val = ws_sum.cell(row=r_sub_excl_cont, column=9, value="=I19+I20")
    c_sub_excl_cont_val.alignment = align_right
    c_sub_excl_cont_val.font = font_subtotal
    c_sub_excl_cont_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_sub_excl_cont, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_sub_excl_cont, column=c).border = border_cell
    ws_sum.row_dimensions[r_sub_excl_cont].height = 23

    # Row 22: Contingencies (10%)
    r_cont = 22
    ws_sum.merge_cells(f"B{r_cont}:H{r_cont}")
    ws_sum.cell(row=r_cont, column=2, value="Contingencies (10%)").alignment = align_right
    ws_sum.cell(row=r_cont, column=2).font = font_data_bold
    c_cont_val = ws_sum.cell(row=r_cont, column=9, value="=I21*'Rates & Parameters'!$F$23")
    c_cont_val.alignment = align_right
    c_cont_val.font = font_data_bold
    c_cont_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_cont, column=c).border = border_cell
    ws_sum.row_dimensions[r_cont].height = 22

    # Row 23: Sub Total: Inclusive of Preliminaries and General & Contingencies (Project Costs A)
    r_sub_incl_all = 23
    ws_sum.merge_cells(f"B{r_sub_incl_all}:H{r_sub_incl_all}")
    ws_sum.cell(row=r_sub_incl_all, column=2, value="Sub Total: Inclusive of Preliminaries and General & Contingencies").alignment = align_right
    ws_sum.cell(row=r_sub_incl_all, column=2).font = font_subtotal
    c_sub_incl_val = ws_sum.cell(row=r_sub_incl_all, column=9, value="=I21+I22")
    c_sub_incl_val.alignment = align_right
    c_sub_incl_val.font = font_subtotal
    c_sub_incl_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_sub_incl_all, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_sub_incl_all, column=c).border = border_subtotal
    ws_sum.row_dimensions[r_sub_incl_all].height = 24

    # Commercial Profit Box & Waterfall
    # Row 25: Profit Percentage
    r_prof_hdr = 25
    ws_sum.merge_cells(f"G{r_prof_hdr}:H{r_prof_hdr}")
    ws_sum[f"G{r_prof_hdr}"] = "Profit Percentage"
    ws_sum[f"G{r_prof_hdr}"].font = font_data_bold
    ws_sum[f"G{r_prof_hdr}"].alignment = align_right
    ws_sum[f"G{r_prof_hdr}"].border = border_cell

    c_prof_pct = ws_sum.cell(row=r_prof_hdr, column=9, value="='Rates & Parameters'!$F$24")
    c_prof_pct.font = font_profit_pct
    c_prof_pct.alignment = align_center
    c_prof_pct.number_format = FMT_PERCENT
    c_prof_pct.border = border_profit_box
    ws_sum.row_dimensions[r_prof_hdr].height = 26

    # Row 26: Monthly Average
    r_monthly = 26
    ws_sum.merge_cells(f"G{r_monthly}:H{r_monthly}")
    ws_sum[f"G{r_monthly}"] = "Monthly"
    ws_sum[f"G{r_monthly}"].font = font_data_bold
    ws_sum[f"G{r_monthly}"].alignment = align_right
    ws_sum[f"G{r_monthly}"].border = border_cell

    c_monthly_val = ws_sum.cell(row=r_monthly, column=9, value="=I23/'Rates & Parameters'!$F$34")
    c_monthly_val.font = font_data
    c_monthly_val.alignment = align_right
    c_monthly_val.number_format = FMT_CURRENCY
    c_monthly_val.border = border_cell
    ws_sum.row_dimensions[r_monthly].height = 22

    # Row 27: Profit
    r_profit = 27
    ws_sum.merge_cells(f"B{r_profit}:H{r_profit}")
    ws_sum[f"B{r_profit}"] = "Profit"
    ws_sum[f"B{r_profit}"].font = font_data_bold
    ws_sum[f"B{r_profit}"].alignment = align_right
    c_prof_val = ws_sum.cell(row=r_profit, column=9, value="=I28*I25")
    c_prof_val.font = font_data_bold
    c_prof_val.alignment = align_right
    c_prof_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_profit, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_profit, column=c).border = border_cell
    ws_sum.row_dimensions[r_profit].height = 22

    # Row 28: Project costs A (Excl. VAT)
    r_proj_cost = 28
    ws_sum.merge_cells(f"B{r_proj_cost}:H{r_proj_cost}")
    ws_sum[f"B{r_proj_cost}"] = "Project costs A (Excl. VAT)"
    ws_sum[f"B{r_proj_cost}"].font = font_data_bold
    ws_sum[f"B{r_proj_cost}"].alignment = align_right
    c_pc_val = ws_sum.cell(row=r_proj_cost, column=9, value="=I23")
    c_pc_val.font = font_data_bold
    c_pc_val.alignment = align_right
    c_pc_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_proj_cost, column=c).border = border_cell
    ws_sum.row_dimensions[r_proj_cost].height = 22

    # Row 29: Sub-Total (B) Tender Price (Excl. VAT)
    r_subtot_fin = 29
    ws_sum.merge_cells(f"B{r_subtot_fin}:H{r_subtot_fin}")
    ws_sum[f"B{r_subtot_fin}"] = "Sub-Total (Tender Price Excl. VAT)"
    ws_sum[f"B{r_subtot_fin}"].font = font_subtotal
    ws_sum[f"B{r_subtot_fin}"].alignment = align_right
    c_st_fin = ws_sum.cell(row=r_subtot_fin, column=9, value="=I28+I27")
    c_st_fin.font = font_subtotal
    c_st_fin.alignment = align_right
    c_st_fin.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_subtot_fin, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_subtot_fin, column=c).border = border_cell
    ws_sum.row_dimensions[r_subtot_fin].height = 23

    # Row 30: 15% VAT
    r_vat = 30
    ws_sum.merge_cells(f"B{r_vat}:H{r_vat}")
    ws_sum[f"B{r_vat}"] = "15% VAT"
    ws_sum[f"B{r_vat}"].font = font_data_bold
    ws_sum[f"B{r_vat}"].alignment = align_right
    c_vat_val = ws_sum.cell(row=r_vat, column=9, value="=I29*'Rates & Parameters'!$F$25")
    c_vat_val.font = font_data_bold
    c_vat_val.alignment = align_right
    c_vat_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_vat, column=c).border = border_cell
    ws_sum.row_dimensions[r_vat].height = 22

    # Row 31: Total including VAT
    r_tot_vat = 31
    ws_sum.merge_cells(f"B{r_tot_vat}:H{r_tot_vat}")
    ws_sum[f"B{r_tot_vat}"] = "Total including VAT"
    ws_sum[f"B{r_tot_vat}"].font = font_data_bold
    ws_sum[f"B{r_tot_vat}"].alignment = align_right
    c_tot_vat_val = ws_sum.cell(row=r_tot_vat, column=9, value="=I29+I30")
    c_tot_vat_val.font = font_data_bold
    c_tot_vat_val.alignment = align_right
    c_tot_vat_val.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_tot_vat, column=c).fill = fill_subtotal
        ws_sum.cell(row=r_tot_vat, column=c).border = border_cell
    ws_sum.row_dimensions[r_tot_vat].height = 23

    # Row 32: Total Project Cost (A) incl VAT
    r_grand_tender = 32
    ws_sum.merge_cells(f"B{r_grand_tender}:H{r_grand_tender}")
    ws_sum[f"B{r_grand_tender}"] = "Total Project Cost (Tender Price Incl. VAT)"
    ws_sum[f"B{r_grand_tender}"].font = font_grandtotal
    ws_sum[f"B{r_grand_tender}"].alignment = align_right
    c_gt_fin = ws_sum.cell(row=r_grand_tender, column=9, value="=I31")
    c_gt_fin.font = font_grandtotal
    c_gt_fin.alignment = align_right
    c_gt_fin.number_format = FMT_CURRENCY
    for c in range(2, 10):
        ws_sum.cell(row=r_grand_tender, column=c).fill = fill_grandtotal
        ws_sum.cell(row=r_grand_tender, column=c).border = border_grandtotal
    ws_sum.row_dimensions[r_grand_tender].height = 27

    # Table 3: Summary per Activity Across All Phases
    r_act_hdr = 35
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

    # Activity Summary rows linking to Activity Summary sheet
    # We will compute the Activity Summary sheet row 16 as the totals row
    act_summary_rows = [
        ("ACT_EST", "Site Establishment", "Phase site establishment, decontamination, mobilization, H&S & staging", "Lot", 9, "=H37/F37", "='Activity Summary'!F16", "=H37/$H$45"),
        ("ACT_01", "1. Steel & Building Structures", "Dismantling, cutting up & removal of all structural steel, headgears & frames", "Structures", 44, "=H38/F38", "='Activity Summary'!G16", "=H38/$H$45"),
        ("ACT_02", "2. Concrete Demolition", "Breakup & stockpile reinforced concrete & brickwork above and ±1.5m below NGL", "m³", 39500, "=H39/F39", "='Activity Summary'!H16", "=H39/$H$45"),
        ("ACT_03", "3. Loading and Hauling", "Loading and hauling clean pulverized rubble & concrete to designated dumping", "R/ton", 116880, "='Rates & Parameters'!$F$13", "='Activity Summary'!I16", "=H40/$H$45"),
        ("ACT_04", "4. Shaping and Levelling", "Final shaping, levelling, compaction and cleaning off rehabilitated area", "m²", 1463268, "='Rates & Parameters'!$F$14", "='Activity Summary'!J16", "=H41/$H$45"),
        ("ACT_SILO", "Specialist Silo Demolition", "Phase 06: Demolition of 2x tall concrete silos & deep raft foundations", "m³", 4200, "='Rates & Parameters'!$F$15", "='Activity Summary'!K16", "=H42/$H$45"),
        ("ACT_WRD", "Waste Rock Dump Processing", "Phase 08: Processing, crushing & screening rock dump for shaft filling", "ton", 1538058, "='Rates & Parameters'!$F$16", "='Activity Summary'!L16", "=H43/$H$45"),
        ("ACT_SHAFT", "Shaft Filling & Capping", "Specialist: Haulage of 1.54M t fill to Main & RV Shafts + 680 m³ capping", "Lot", 1, "=H44/F44", "='Activity Summary'!M16", "=H44/$H$45")
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

    # Grand Total Activity Summary Row (Row 45)
    r_act_tot = 45
    ws_sum.merge_cells(f"B{r_act_tot}:G{r_act_tot}")
    ws_sum[f"B{r_act_tot}"] = "TOTAL DIRECT ACTIVITY COST (RECONCILED TO PHASE DIRECT TOTAL)"
    ws_sum[f"B{r_act_tot}"].font = font_grandtotal
    ws_sum[f"B{r_act_tot}"].alignment = align_right

    c_act_tot_val = ws_sum.cell(row=r_act_tot, column=8, value=f"=SUM(H37:H44)")
    c_act_tot_val.font = font_grandtotal
    c_act_tot_val.alignment = align_right
    c_act_tot_val.number_format = FMT_CURRENCY

    c_act_pct_tot = ws_sum.cell(row=r_act_tot, column=9, value=f"=SUM(I37:I44)")
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
    ws_sum.column_dimensions["D"].width = 46
    ws_sum.column_dimensions["E"].width = 16
    ws_sum.column_dimensions["F"].width = 14
    ws_sum.column_dimensions["G"].width = 16
    ws_sum.column_dimensions["H"].width = 22
    ws_sum.column_dimensions["I"].width = 22

    # =========================================================================
    # SHEET 4: ACTIVITY SUMMARY (CROSS-PHASE BREAKDOWN MATRIX)
    # =========================================================================
    ws_act = wb.create_sheet(title="Activity Summary")
    ws_act.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_act.merge_cells("B2:N2")
    ws_act["B2"] = "CAPM TAU MINE #10 - CROSS-PHASE ACTIVITY BREAKDOWN MATRIX"
    ws_act["B2"].font = font_main_title
    ws_act["B2"].fill = fill_main_title
    ws_act["B2"].alignment = align_title
    ws_act.row_dimensions[2].height = 32

    ws_act.merge_cells("B3:N3")
    ws_act["B3"] = "Direct Cost Distribution per Activity Across All Phases (Steel, Concrete, Load & Haul, Shaping & Levelling, Waste Rock, Shafts)"
    ws_act["B3"].font = font_sub_title
    ws_act["B3"].fill = fill_sec_hdr
    ws_act["B3"].alignment = align_center
    ws_act.row_dimensions[3].height = 20

    act_matrix_headers = [
        "Item", "Mine Phase Description", "Area (m²)", "Duration (Mos)",
        "Site Est. (Lot)", "1. Steel & Struct. (R)", "2. Concrete (R)",
        "3. Load & Haul (R)", "4. Shaping & Level (R)", "Specialist Silos (R)",
        "Waste Rock Proc. (R)", "Shaft Filling & Cap (R)", "Phase Total (R)"
    ]
    for col_idx, h in enumerate(act_matrix_headers, start=2):
        cell = ws_act.cell(row=5, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center if col_idx in [2, 4, 5] else (align_right if col_idx >= 6 else align_left)
        cell.border = border_cell
    ws_act.row_dimensions[5].height = 26

    # Matrix Phase Links directly from Master BOQ Breakdown
    # Let's verify each phase's item rows in Master BOQ Breakdown
    # Phase 1: rows 6..10, subtotal 11 -> items: 6=Est, 7=Steel, 8=Conc, 9=Haul, 10=Shape
    # Phase 2: rows 12..16, subtotal 17 -> items: 12=Est, 13=Steel, 14=Conc, 15=Haul, 16=Shape
    # Phase 3: rows 18..22, subtotal 23 -> items: 18=Est, 19=Steel, 20=Conc, 21=Haul, 22=Shape
    # Phase 4: rows 24..28, subtotal 29 -> items: 24=Est, 25=Steel, 26=Conc, 27=Haul, 28=Shape
    # Phase 5: rows 30..34, subtotal 35 -> items: 30=Est, 31=Steel, 32=Conc, 33=Haul, 34=Shape
    # Phase 6: rows 36..43, subtotal 44 -> items: 36=Est, 37+38+39=Steel, 41=Conc, 42=Haul, 43=Shape, 40=Silos
    # Phase 7: rows 45..49, subtotal 50 -> items: 45=Est, 46=Steel, 47=Conc, 48=Haul, 49=Shape
    # Phase 8: rows 51..56, subtotal 57 -> items: 51=Est, 52=Steel, 53=Conc, 54=WRD, 55=Haul, 56=Shape
    # Phase 9: rows 58..62, subtotal 63 -> items: 58=Est, 59+60+61+62=Shafts
    matrix_phase_links = [
        (1, "Phase 01: Salvage Yard & Slabs", 178860.00, 3, "='Master BOQ Breakdown'!G6", "='Master BOQ Breakdown'!G7", "='Master BOQ Breakdown'!G8", "='Master BOQ Breakdown'!G9", "='Master BOQ Breakdown'!G10", 0, 0, 0, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[1]}"),
        (2, "Phase 02: Admin & Cooling Towers", 102294.00, 4, "='Master BOQ Breakdown'!G12", "='Master BOQ Breakdown'!G13", "='Master BOQ Breakdown'!G14", "='Master BOQ Breakdown'!G15", "='Master BOQ Breakdown'!G16", 0, 0, 0, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[2]}"),
        (3, "Phase 03: Kitchen & Offices Area", 141802.00, 4, "='Master BOQ Breakdown'!G18", "='Master BOQ Breakdown'!G19", "='Master BOQ Breakdown'!G20", "='Master BOQ Breakdown'!G21", "='Master BOQ Breakdown'!G22", 0, 0, 0, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[3]}"),
        (4, "Phase 04: Hostel Complex Area", 12763.00, 4, "='Master BOQ Breakdown'!G24", "='Master BOQ Breakdown'!G25", "='Master BOQ Breakdown'!G26", "='Master BOQ Breakdown'!G27", "='Master BOQ Breakdown'!G28", 0, 0, 0, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[4]}"),
        (5, "Phase 05: Workshops & Substation", 119515.00, 3, "='Master BOQ Breakdown'!G30", "='Master BOQ Breakdown'!G31", "='Master BOQ Breakdown'!G32", "='Master BOQ Breakdown'!G33", "='Master BOQ Breakdown'!G34", 0, 0, 0, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[5]}"),
        (6, "Phase 06: Shafts, Silos & Headgears", 91628.00, 6, "='Master BOQ Breakdown'!G36", "=SUM('Master BOQ Breakdown'!G37:G39)", "='Master BOQ Breakdown'!G41", "='Master BOQ Breakdown'!G42", "='Master BOQ Breakdown'!G43", "='Master BOQ Breakdown'!G40", 0, 0, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[6]}"),
        (7, "Phase 07: Dams & Water Reservoirs", 264040.00, 2, "='Master BOQ Breakdown'!G45", "='Master BOQ Breakdown'!G46", "='Master BOQ Breakdown'!G47", "='Master BOQ Breakdown'!G48", "='Master BOQ Breakdown'!G49", 0, 0, 0, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[7]}"),
        (8, "Phase 08: Waste Rock Dump Process", 552366.00, 12, "='Master BOQ Breakdown'!G51", "='Master BOQ Breakdown'!G52", "='Master BOQ Breakdown'!G53", "='Master BOQ Breakdown'!G55", "='Master BOQ Breakdown'!G56", 0, "='Master BOQ Breakdown'!G54", 0, f"='Master BOQ Breakdown'!G{phase_subtotal_rows[8]}"),
        (9, "Specialist: Shaft Filling & Capping", 0.0, 6, "='Master BOQ Breakdown'!G58", 0, 0, 0, 0, 0, 0, "=SUM('Master BOQ Breakdown'!G59:G62)", f"='Master BOQ Breakdown'!G{phase_subtotal_rows[9]}")
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
            if col_i == 14:
                cell.font = font_data_bold
            else:
                cell.font = font_data

        fill_row = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)
        for c in range(2, 15):
            ws_act.cell(row=idx, column=c).border = border_cell
            if fill_row.fill_type:
                ws_act.cell(row=idx, column=c).fill = fill_row
        ws_act.row_dimensions[idx].height = 22

    # Grand Totals Row for Matrix (Row 15) -> Let's make it Row 16 to leave clean spacing
    r_mat_tot = 16
    ws_act.merge_cells(f"B{r_mat_tot}:C{r_mat_tot}")
    ws_act[f"B{r_mat_tot}"] = "TOTAL DIRECT PROJECT COST"
    ws_act[f"B{r_mat_tot}"].font = font_grandtotal
    ws_act[f"B{r_mat_tot}"].alignment = align_right

    c_tot_area = ws_act.cell(row=r_mat_tot, column=4, value="=SUM(D6:D14)")
    c_tot_area.alignment = align_right
    c_tot_area.font = font_grandtotal
    c_tot_area.number_format = FMT_QTY_DEC

    c_tot_dur = ws_act.cell(row=r_mat_tot, column=5, value="=SUM(E6:E14)")
    c_tot_dur.alignment = align_center
    c_tot_dur.font = font_grandtotal
    c_tot_dur.number_format = FMT_QTY

    # Columns F through N
    for col_i, col_letter in enumerate(["F", "G", "H", "I", "J", "K", "L", "M", "N"], start=6):
        cell = ws_act.cell(row=r_mat_tot, column=col_i, value=f"=SUM({col_letter}6:{col_letter}14)")
        cell.alignment = align_right
        cell.font = font_grandtotal
        cell.number_format = FMT_CURRENCY

    for c in range(2, 15):
        ws_act.cell(row=r_mat_tot, column=c).fill = fill_grandtotal
        ws_act.cell(row=r_mat_tot, column=c).border = border_grandtotal
    ws_act.row_dimensions[r_mat_tot].height = 25

    # Percentage Contribution Row (Row 17)
    r_mat_pct = 17
    ws_act.merge_cells(f"B{r_mat_pct}:E{r_mat_pct}")
    ws_act[f"B{r_mat_pct}"] = "% Contribution to Total Direct Spend"
    ws_act[f"B{r_mat_pct}"].font = font_subtotal
    ws_act[f"B{r_mat_pct}"].alignment = align_right

    for col_i, col_letter in enumerate(["F", "G", "H", "I", "J", "K", "L", "M"], start=6):
        cell = ws_act.cell(row=r_mat_pct, column=col_i, value=f"={col_letter}{r_mat_tot}/$N${r_mat_tot}")
        cell.alignment = align_right
        cell.font = font_subtotal
        cell.number_format = FMT_PERCENT
        cell.border = border_subtotal
        cell.fill = fill_subtotal

    cell_n_pct = ws_act.cell(row=r_mat_pct, column=14, value=f"=SUM(F{r_mat_pct}:M{r_mat_pct})")
    cell_n_pct.alignment = align_right
    cell_n_pct.font = font_subtotal
    cell_n_pct.number_format = FMT_PERCENT
    cell_n_pct.border = border_subtotal
    cell_n_pct.fill = fill_subtotal

    for c in range(2, 6):
        ws_act.cell(row=r_mat_pct, column=c).fill = fill_subtotal
        ws_act.cell(row=r_mat_pct, column=c).border = border_subtotal
    ws_act.row_dimensions[r_mat_pct].height = 22

    # Column widths for Activity Summary
    ws_act.column_dimensions["A"].width = 3
    ws_act.column_dimensions["B"].width = 8
    ws_act.column_dimensions["C"].width = 32
    ws_act.column_dimensions["D"].width = 16
    ws_act.column_dimensions["E"].width = 15
    for cl in ["F", "G", "H", "I", "J", "K", "L", "M", "N"]:
        ws_act.column_dimensions[cl].width = 20

    # =========================================================================
    # DEDICATED PHASE SHEETS (9 WORK PACKAGES: PHASES 01 TO 08 + SHAFTS)
    # =========================================================================
    # Sheet configs: (item_num, title, sheet_tab_name, boq_start_row, boq_end_row, boq_subtotal_row)
    phase_sheet_configs = [
        (1, "PHASE 01: SALVAGE YARD & SLABS AREA", "Phase 1 - Salvage & Slabs", 178860.00, 3, 6, 10800),
        (2, "PHASE 02: ADMIN & COOLING TOWERS AREA", "Phase 2 - Admin & Cooling", 102294.00, 4, 8, 14880),
        (3, "PHASE 03: KITCHEN & OFFICES AREA", "Phase 3 - Kitchen Complex", 141802.00, 4, 5, 13920),
        (4, "PHASE 04: HOSTEL COMPLEX AREA", "Phase 4 - Hostel Area", 12763.00, 4, 6, 8400),
        (5, "PHASE 05: WORKSHOPS & SUBSTATION AREA", "Phase 5 - Workshops & Sub", 119515.00, 3, 7, 12480),
        (6, "PHASE 06: SHAFTS, SILOS & HEADGEARS", "Phase 6 - Shafts & Headgears", 91628.00, 6, 10, 30720),
        (7, "PHASE 07: DAMS & WATER RESERVOIRS AREA", "Phase 7 - Dams & Reservoir", 264040.00, 2, 4, 18000),
        (8, "PHASE 08: WASTE ROCK DUMP PROCESSING", "Phase 8 - Waste Rock Dump", 552366.00, 12, 2, 7680),
        (9, "SPECIALIST: SHAFT FILLING & CAPPING", "Shaft Filling & Capping", 0.0, 6, 0, 1538058)
    ]

    for p_cfg in phase_sheet_configs:
        p_num, p_full_title, p_sheet_name, p_area, p_dur, p_struc, p_rubble = p_cfg
        ws_p = wb.create_sheet(title=p_sheet_name)
        ws_p.views.sheetView[0].showGridLines = True

        # Banner
        ws_p.merge_cells("B2:G2")
        ws_p["B2"] = f"CAPM TAU MINE #10 - {p_full_title}"
        ws_p["B2"].font = font_main_title
        ws_p["B2"].fill = fill_main_title
        ws_p["B2"].alignment = align_title
        ws_p.row_dimensions[2].height = 32

        ws_p.merge_cells("B3:G3")
        if p_area > 0:
            ws_p["B3"] = f"Footprint: {p_area:,.2f} m² | Duration: ±{p_dur} Months | Clearance Depth: ±1.5m below NGL | Demolition, Haul & Surface Shaping"
        else:
            ws_p["B3"] = f"Specialist Scope: 1,538,058 t Shaft Backfill | 680 m³ Reinforced Concrete Capping | ±{p_dur} Months Execution"
        ws_p["B3"].font = font_sub_title
        ws_p["B3"].fill = fill_sec_hdr
        ws_p["B3"].alignment = align_center
        ws_p.row_dimensions[3].height = 20

        # KPI Block
        ws_p.merge_cells("B5:C5")
        ws_p["B5"] = p_area
        ws_p["B5"].font = font_kpi_num
        ws_p["B5"].alignment = align_center
        ws_p["B5"].fill = fill_kpi
        ws_p["B5"].number_format = '#,##0.00" m²"' if p_area > 0 else '"Deep Shafts"'

        ws_p.merge_cells("B6:C6")
        ws_p["B6"] = "SURFACE FOOTPRINT"
        ws_p["B6"].font = font_kpi_lbl
        ws_p["B6"].alignment = align_center
        ws_p["B6"].fill = fill_kpi

        ws_p.merge_cells("D5:D5")
        ws_p["D5"] = p_dur
        ws_p["D5"].font = font_kpi_num
        ws_p["D5"].alignment = align_center
        ws_p["D5"].fill = fill_kpi
        ws_p["D5"].number_format = '#,##0" Months"'

        ws_p.merge_cells("D6:D6")
        ws_p["D6"] = "DURATION"
        ws_p["D6"].font = font_kpi_lbl
        ws_p["D6"].alignment = align_center
        ws_p["D6"].fill = fill_kpi

        ws_p.merge_cells("E5:E5")
        ws_p["E5"] = p_rubble
        ws_p["E5"].font = font_kpi_num
        ws_p["E5"].alignment = align_center
        ws_p["E5"].fill = fill_kpi
        ws_p["E5"].number_format = '#,##0" t"'

        ws_p.merge_cells("E6:E6")
        ws_p["E6"] = "MATERIAL TONNAGE"
        ws_p["E6"].font = font_kpi_lbl
        ws_p["E6"].alignment = align_center
        ws_p["E6"].fill = fill_kpi

        ws_p.merge_cells("F5:G5")
        b_sub_row = phase_subtotal_rows[p_num]
        ws_p["F5"] = f"='Master BOQ Breakdown'!G{b_sub_row}"
        ws_p["F5"].font = font_kpi_num
        ws_p["F5"].alignment = align_center
        ws_p["F5"].fill = fill_kpi
        ws_p["F5"].number_format = FMT_CURRENCY

        ws_p.merge_cells("F6:G6")
        ws_p["F6"] = "DIRECT PACKAGE BUDGET"
        ws_p["F6"].font = font_kpi_lbl
        ws_p["F6"].alignment = align_center
        ws_p["F6"].fill = fill_kpi

        for cell_ref in ["B5", "B6", "C5", "C6", "D5", "D6", "E5", "E6", "F5", "F6", "G5", "G6"]:
            ws_p[cell_ref].border = border_kpi
        ws_p.row_dimensions[5].height = 24
        ws_p.row_dimensions[6].height = 18

        # BOQ Table Headers
        r_p_boq = 8
        p_headers = ["Item", "Technical Activity & Line Item Scope", "UOM", "Quantity", "Unit Rate (R)", "Total Direct Cost (R)"]
        for col_idx, h in enumerate(p_headers, start=2):
            cell = ws_p.cell(row=r_p_boq, column=col_idx, value=h)
            cell.font = font_tbl_hdr
            cell.fill = fill_main_title
            cell.alignment = align_center if col_idx in [2, 4] else (align_right if col_idx in [5, 6, 7] else align_left)
            cell.border = border_cell
        ws_p.row_dimensions[r_p_boq].height = 25

        # Populate Items by dynamic reference to Master BOQ Breakdown
        b_start, b_end = phase_start_end_rows[p_num]
        num_p_items = b_end - b_start + 1

        for itm_i in range(num_p_items):
            cur_p_r = r_p_boq + 1 + itm_i
            cur_b_r = b_start + itm_i

            ws_p.cell(row=cur_p_r, column=2, value=f"{p_num}.{itm_i+1}").alignment = align_center
            ws_p.cell(row=cur_p_r, column=2).font = font_data

            c_d = ws_p.cell(row=cur_p_r, column=3, value=f"='Master BOQ Breakdown'!C{cur_b_r}")
            c_d.alignment = align_left_wrap
            c_d.font = font_data

            c_u = ws_p.cell(row=cur_p_r, column=4, value=f"='Master BOQ Breakdown'!D{cur_b_r}")
            c_u.alignment = align_center
            c_u.font = font_data

            c_q = ws_p.cell(row=cur_p_r, column=5, value=f"='Master BOQ Breakdown'!E{cur_b_r}")
            c_q.alignment = align_right
            c_q.font = font_data
            c_q.number_format = FMT_QTY_DEC if p_num == 9 and itm_i in [1, 3] else FMT_QTY

            c_r = ws_p.cell(row=cur_p_r, column=6, value=f"='Master BOQ Breakdown'!F{cur_b_r}")
            c_r.alignment = align_right
            c_r.font = font_data
            c_r.number_format = FMT_CURRENCY

            c_t = ws_p.cell(row=cur_p_r, column=7, value=f"=E{cur_p_r}*F{cur_p_r}")
            c_t.alignment = align_right
            c_t.font = font_data_bold
            c_t.number_format = FMT_CURRENCY

            fill_row = fill_zebra if itm_i % 2 == 1 else PatternFill(fill_type=None)
            for c in range(2, 8):
                ws_p.cell(row=cur_p_r, column=c).border = border_cell
                if fill_row.fill_type:
                    ws_p.cell(row=cur_p_r, column=c).fill = fill_row
            ws_p.row_dimensions[cur_p_r].height = 22

        # Sub-Total Row
        r_p_sub = r_p_boq + 1 + num_p_items
        ws_p.merge_cells(f"B{r_p_sub}:F{r_p_sub}")
        ws_p.cell(row=r_p_sub, column=2, value=f"SUB-TOTAL: {p_full_title}").alignment = align_right
        ws_p.cell(row=r_p_sub, column=2).font = font_subtotal

        c_p_sub_val = ws_p.cell(row=r_p_sub, column=7, value=f"=SUM(G{r_p_boq+1}:G{r_p_sub-1})")
        c_p_sub_val.alignment = align_right
        c_p_sub_val.font = font_subtotal
        c_p_sub_val.number_format = FMT_CURRENCY

        for c in range(2, 8):
            ws_p.cell(row=r_p_sub, column=c).fill = fill_subtotal
            ws_p.cell(row=r_p_sub, column=c).border = border_subtotal
        ws_p.row_dimensions[r_p_sub].height = 24

        # Reconciliation Note
        r_p_note = r_p_sub + 2
        ws_p.merge_cells(f"B{r_p_note}:G{r_p_note}")
        ws_p.cell(row=r_p_note, column=2, value=f"Commercial Note: Direct package costs feed directly into Master BOQ Breakdown (Row {b_sub_row}) and Executive Summary Table 1.").font = font_note
        ws_p.cell(row=r_p_note, column=2).alignment = align_left

        # Column widths
        ws_p.column_dimensions["A"].width = 3
        ws_p.column_dimensions["B"].width = 10
        ws_p.column_dimensions["C"].width = 62
        ws_p.column_dimensions["D"].width = 14
        ws_p.column_dimensions["E"].width = 18
        ws_p.column_dimensions["F"].width = 20
        ws_p.column_dimensions["G"].width = 24

    # =========================================================================
    # SHEET 5: SCHEDULE & CASHFLOW (24-MONTH TIMELINE & S-CURVE FORECAST)
    # =========================================================================
    ws_sched = wb.create_sheet(title="Schedule & Cashflow")
    ws_sched.views.sheetView[0].showGridLines = True

    ws_sched.merge_cells("B2:AB2")
    ws_sched["B2"] = "CAPM TAU MINE #10 - 24-MONTH PROJECT TIMELINE & CASH FLOW FORECAST"
    ws_sched["B2"].font = font_main_title
    ws_sched["B2"].fill = fill_main_title
    ws_sched["B2"].alignment = align_title
    ws_sched.row_dimensions[2].height = 32

    ws_sched.merge_cells("B3:AB3")
    ws_sched["B3"] = "Phased Execution Plan, Monthly Billing Distribution, Markups Allocation & Cumulative S-Curve (Phases 01 to 08 + Shaft Filling)"
    ws_sched["B3"].font = font_sub_title
    ws_sched["B3"].fill = fill_sec_hdr
    ws_sched["B3"].alignment = align_center
    ws_sched.row_dimensions[3].height = 20

    sched_hdrs = ["Item", "Mine Phase / Contract Package", "Direct Cost (R)"] + [f"M{m:02d}" for m in range(1, 25)]
    for col_idx, h in enumerate(sched_hdrs, start=2):
        cell = ws_sched.cell(row=5, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_main_title
        cell.alignment = align_center
        cell.border = border_cell
    ws_sched.row_dimensions[5].height = 25

    # 24-Month Phased Distribution Strategy:
    # Phase 1: M01 - M03 (3 mos: 33.33%, 33.33%, 33.34%)
    # Phase 2: M02 - M05 (4 mos: 25%, 25%, 25%, 25%)
    # Phase 3: M04 - M07 (4 mos: 25%, 25%, 25%, 25%)
    # Phase 4: M06 - M09 (4 mos: 25%, 25%, 25%, 25%)
    # Phase 5: M08 - M10 (3 mos: 33.33%, 33.33%, 33.34%)
    # Phase 6: M10 - M15 (6 mos: 16.66% * 5 + 16.70%)
    # Phase 7: M14 - M15 (2 mos: 50%, 50%)
    # Phase 8: M11 - M22 (12 mos: 8.33% * 11 + 8.37%)
    # Phase 9: M17 - M22 (6 mos: 16.66% * 5 + 16.70%) [Shaft filling uses processed rock from Phase 8!]
    # Months 23-24: Site final shaping, contouring, environmental auditing & client handover
    sched_schedule_data = [
        (1, "Phase 01: Salvage Yard & Slabs Area", "='Executive Summary & Financials'!I10", {1: 0.3333, 2: 0.3333, 3: 0.3334}),
        (2, "Phase 02: Admin & Cooling Towers Area", "='Executive Summary & Financials'!I11", {2: 0.25, 3: 0.25, 4: 0.25, 5: 0.25}),
        (3, "Phase 03: Kitchen & Offices Area", "='Executive Summary & Financials'!I12", {4: 0.25, 5: 0.25, 6: 0.25, 7: 0.25}),
        (4, "Phase 04: Hostel Complex Area", "='Executive Summary & Financials'!I13", {6: 0.25, 7: 0.25, 8: 0.25, 9: 0.25}),
        (5, "Phase 05: Workshops & Substation Area", "='Executive Summary & Financials'!I14", {8: 0.3333, 9: 0.3333, 10: 0.3334}),
        (6, "Phase 06: Shafts, Silos & Headgears", "='Executive Summary & Financials'!I15", {10: 0.1666, 11: 0.1666, 12: 0.1667, 13: 0.1667, 14: 0.1667, 15: 0.1667}),
        (7, "Phase 07: Dams & Water Reservoirs", "='Executive Summary & Financials'!I16", {14: 0.50, 15: 0.50}),
        (8, "Phase 08: Waste Rock Dump Processing", "='Executive Summary & Financials'!I17", {m: (0.0837 if m == 22 else 0.0833) for m in range(11, 23)}),
        (9, "Specialist: Shaft Filling & Capping", "='Executive Summary & Financials'!I18", {17: 0.1666, 18: 0.1666, 19: 0.1667, 20: 0.1667, 21: 0.1667, 22: 0.1667})
    ]

    for idx, s in enumerate(sched_schedule_data, start=6):
        item_i, name_s, tot_fml, dist_map = s
        ws_sched.cell(row=idx, column=2, value=item_i).alignment = align_center
        ws_sched.cell(row=idx, column=3, value=name_s).alignment = align_left
        
        c_tot = ws_sched.cell(row=idx, column=4, value=tot_fml)
        c_tot.alignment = align_right
        c_tot.font = font_data_bold
        c_tot.number_format = FMT_CURRENCY

        for m in range(1, 25):
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

    # Monthly Direct Cost Row (Row 15)
    r_mo_dir = 15
    ws_sched.merge_cells(f"B{r_mo_dir}:C{r_mo_dir}")
    ws_sched[f"B{r_mo_dir}"] = "MONTHLY DIRECT PROJECT SPEND"
    ws_sched[f"B{r_mo_dir}"].font = font_subtotal
    ws_sched[f"B{r_mo_dir}"].alignment = align_right

    c_dir_tot = ws_sched.cell(row=r_mo_dir, column=4, value="=SUM(D6:D14)")
    c_dir_tot.alignment = align_right
    c_dir_tot.font = font_subtotal
    c_dir_tot.number_format = FMT_CURRENCY

    for m in range(1, 25):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_dir, column=col_m, value=f"=SUM({col_letter}6:{col_letter}14)")
        cell.alignment = align_right
        cell.font = font_subtotal
        cell.number_format = FMT_CURRENCY

    for c in range(2, 29):
        ws_sched.cell(row=r_mo_dir, column=c).fill = fill_subtotal
        ws_sched.cell(row=r_mo_dir, column=c).border = border_subtotal
    ws_sched.row_dimensions[r_mo_dir].height = 23

    # Monthly Commercial Markups & Cashflow Rows
    # Row 16: Preliminaries and General (10%)
    r_mo_pg = 16
    ws_sched.merge_cells(f"B{r_mo_pg}:C{r_mo_pg}")
    ws_sched[f"B{r_mo_pg}"] = "Preliminaries and General (10%)"
    ws_sched[f"B{r_mo_pg}"].font = font_data_bold
    ws_sched[f"B{r_mo_pg}"].alignment = align_right

    c_mo_pg_tot = ws_sched.cell(row=r_mo_pg, column=4, value=f"=D{r_mo_dir}*'Rates & Parameters'!$F$22")
    c_mo_pg_tot.alignment = align_right
    c_mo_pg_tot.font = font_data_bold
    c_mo_pg_tot.number_format = FMT_CURRENCY

    for m in range(1, 25):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_pg, column=col_m, value=f"={col_letter}{r_mo_dir}*'Rates & Parameters'!$F$22")
        cell.alignment = align_right
        cell.font = font_data
        cell.number_format = FMT_CURRENCY
        cell.border = border_cell

    for c in [2, 3, 4]:
        ws_sched.cell(row=r_mo_pg, column=c).border = border_cell
    ws_sched.row_dimensions[r_mo_pg].height = 22

    # Row 17: Contingencies (10%)
    r_mo_cont = 17
    ws_sched.merge_cells(f"B{r_mo_cont}:C{r_mo_cont}")
    ws_sched[f"B{r_mo_cont}"] = "Contingencies Allowance (10%)"
    ws_sched[f"B{r_mo_cont}"].font = font_data_bold
    ws_sched[f"B{r_mo_cont}"].alignment = align_right

    c_mo_cont_tot = ws_sched.cell(row=r_mo_cont, column=4, value=f"=(D{r_mo_dir}+D{r_mo_pg})*'Rates & Parameters'!$F$23")
    c_mo_cont_tot.alignment = align_right
    c_mo_cont_tot.font = font_data_bold
    c_mo_cont_tot.number_format = FMT_CURRENCY

    for m in range(1, 25):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_cont, column=col_m, value=f"=({col_letter}{r_mo_dir}+{col_letter}{r_mo_pg})*'Rates & Parameters'!$F$23")
        cell.alignment = align_right
        cell.font = font_data
        cell.number_format = FMT_CURRENCY
        cell.border = border_cell

    for c in [2, 3, 4]:
        ws_sched.cell(row=r_mo_cont, column=c).border = border_cell
    ws_sched.row_dimensions[r_mo_cont].height = 22

    # Row 18: Project Costs A (Excl. Profit & VAT)
    r_mo_cost_a = 18
    ws_sched.merge_cells(f"B{r_mo_cost_a}:C{r_mo_cost_a}")
    ws_sched[f"B{r_mo_cost_a}"] = "Total Project Costs A (Excl. Profit & VAT)"
    ws_sched[f"B{r_mo_cost_a}"].font = font_subtotal
    ws_sched[f"B{r_mo_cost_a}"].alignment = align_right

    c_mo_a_tot = ws_sched.cell(row=r_mo_cost_a, column=4, value=f"=D{r_mo_dir}+D{r_mo_pg}+D{r_mo_cont}")
    c_mo_a_tot.alignment = align_right
    c_mo_a_tot.font = font_subtotal
    c_mo_a_tot.number_format = FMT_CURRENCY

    for m in range(1, 25):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_cost_a, column=col_m, value=f"={col_letter}{r_mo_dir}+{col_letter}{r_mo_pg}+{col_letter}{r_mo_cont}")
        cell.alignment = align_right
        cell.font = font_subtotal
        cell.number_format = FMT_CURRENCY
        cell.border = border_subtotal
        cell.fill = fill_subtotal

    for c in [2, 3, 4]:
        ws_sched.cell(row=r_mo_cost_a, column=c).fill = fill_subtotal
        ws_sched.cell(row=r_mo_cost_a, column=c).border = border_subtotal
    ws_sched.row_dimensions[r_mo_cost_a].height = 23

    # Row 19: Contractor Profit (5%)
    r_mo_prof = 19
    ws_sched.merge_cells(f"B{r_mo_prof}:C{r_mo_prof}")
    ws_sched[f"B{r_mo_prof}"] = "Contractor Net Profit (5%)"
    ws_sched[f"B{r_mo_prof}"].font = font_data_bold
    ws_sched[f"B{r_mo_prof}"].alignment = align_right

    c_mo_prof_tot = ws_sched.cell(row=r_mo_prof, column=4, value=f"=D{r_mo_cost_a}*'Rates & Parameters'!$F$24")
    c_mo_prof_tot.alignment = align_right
    c_mo_prof_tot.font = font_data_bold
    c_mo_prof_tot.number_format = FMT_CURRENCY

    for m in range(1, 25):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_prof, column=col_m, value=f"={col_letter}{r_mo_cost_a}*'Rates & Parameters'!$F$24")
        cell.alignment = align_right
        cell.font = font_data
        cell.number_format = FMT_CURRENCY
        cell.border = border_cell

    for c in [2, 3, 4]:
        ws_sched.cell(row=r_mo_prof, column=c).border = border_cell
    ws_sched.row_dimensions[r_mo_prof].height = 22

    # Row 20: Sub-Total Tender Price (Excl. VAT)
    r_mo_sub_b = 20
    ws_sched.merge_cells(f"B{r_mo_sub_b}:C{r_mo_sub_b}")
    ws_sched[f"B{r_mo_sub_b}"] = "Sub-Total Tender Billing (Excl. VAT)"
    ws_sched[f"B{r_mo_sub_b}"].font = font_subtotal
    ws_sched[f"B{r_mo_sub_b}"].alignment = align_right

    c_mo_b_tot = ws_sched.cell(row=r_mo_sub_b, column=4, value=f"=D{r_mo_cost_a}+D{r_mo_prof}")
    c_mo_b_tot.alignment = align_right
    c_mo_b_tot.font = font_subtotal
    c_mo_b_tot.number_format = FMT_CURRENCY

    for m in range(1, 25):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_sub_b, column=col_m, value=f"={col_letter}{r_mo_cost_a}+{col_letter}{r_mo_prof}")
        cell.alignment = align_right
        cell.font = font_subtotal
        cell.number_format = FMT_CURRENCY
        cell.border = border_subtotal
        cell.fill = fill_subtotal

    for c in [2, 3, 4]:
        ws_sched.cell(row=r_mo_sub_b, column=c).fill = fill_subtotal
        ws_sched.cell(row=r_mo_sub_b, column=c).border = border_subtotal
    ws_sched.row_dimensions[r_mo_sub_b].height = 23

    # Row 21: 15% VAT
    r_mo_vat = 21
    ws_sched.merge_cells(f"B{r_mo_vat}:C{r_mo_vat}")
    ws_sched[f"B{r_mo_vat}"] = "Value Added Tax (15% SARS VAT)"
    ws_sched[f"B{r_mo_vat}"].font = font_data_bold
    ws_sched[f"B{r_mo_vat}"].alignment = align_right

    c_mo_vat_tot = ws_sched.cell(row=r_mo_vat, column=4, value=f"=D{r_mo_sub_b}*'Rates & Parameters'!$F$25")
    c_mo_vat_tot.alignment = align_right
    c_mo_vat_tot.font = font_data_bold
    c_mo_vat_tot.number_format = FMT_CURRENCY

    for m in range(1, 25):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_vat, column=col_m, value=f"={col_letter}{r_mo_sub_b}*'Rates & Parameters'!$F$25")
        cell.alignment = align_right
        cell.font = font_data
        cell.number_format = FMT_CURRENCY
        cell.border = border_cell

    for c in [2, 3, 4]:
        ws_sched.cell(row=r_mo_vat, column=c).border = border_cell
    ws_sched.row_dimensions[r_mo_vat].height = 22

    # Row 22: Monthly Total Billing (Incl. VAT)
    r_mo_gross = 22
    ws_sched.merge_cells(f"B{r_mo_gross}:C{r_mo_gross}")
    ws_sched[f"B{r_mo_gross}"] = "MONTHLY GROSS INVOICING (INCL. VAT)"
    ws_sched[f"B{r_mo_gross}"].font = font_grandtotal
    ws_sched[f"B{r_mo_gross}"].alignment = align_right

    c_mo_gross_tot = ws_sched.cell(row=r_mo_gross, column=4, value=f"=D{r_mo_sub_b}+D{r_mo_vat}")
    c_mo_gross_tot.alignment = align_right
    c_mo_gross_tot.font = font_grandtotal
    c_mo_gross_tot.number_format = FMT_CURRENCY

    for m in range(1, 25):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_gross, column=col_m, value=f"={col_letter}{r_mo_sub_b}+{col_letter}{r_mo_vat}")
        cell.alignment = align_right
        cell.font = font_grandtotal
        cell.number_format = FMT_CURRENCY

    for c in range(2, 29):
        ws_sched.cell(row=r_mo_gross, column=c).fill = fill_grandtotal
        ws_sched.cell(row=r_mo_gross, column=c).border = border_grandtotal
    ws_sched.row_dimensions[r_mo_gross].height = 26

    # Row 23: Cumulative Invoiced Cash Flow
    r_mo_cum = 23
    ws_sched.merge_cells(f"B{r_mo_cum}:C{r_mo_cum}")
    ws_sched[f"B{r_mo_cum}"] = "CUMULATIVE CASH FLOW INVOICED (INCL. VAT)"
    ws_sched[f"B{r_mo_cum}"].font = font_subtotal
    ws_sched[f"B{r_mo_cum}"].alignment = align_right

    c_cum_tot = ws_sched.cell(row=r_mo_cum, column=4, value=f"=AB{r_mo_cum}")
    c_cum_tot.alignment = align_right
    c_cum_tot.font = font_subtotal
    c_cum_tot.number_format = FMT_CURRENCY

    # Column E is M01: =E22
    ws_sched.cell(row=r_mo_cum, column=5, value="=E22").alignment = align_right
    ws_sched.cell(row=r_mo_cum, column=5).font = font_subtotal
    ws_sched.cell(row=r_mo_cum, column=5).number_format = FMT_CURRENCY

    for m in range(2, 25):
        col_m = 4 + m
        col_prev = get_column_letter(col_m - 1)
        col_cur = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_cum, column=col_m, value=f"={col_prev}{r_mo_cum}+{col_cur}{r_mo_gross}")
        cell.alignment = align_right
        cell.font = font_subtotal
        cell.number_format = FMT_CURRENCY

    for c in range(2, 29):
        ws_sched.cell(row=r_mo_cum, column=c).fill = fill_subtotal
        ws_sched.cell(row=r_mo_cum, column=c).border = border_subtotal
    ws_sched.row_dimensions[r_mo_cum].height = 23

    # Row 24: Cumulative S-Curve Progress (%)
    r_mo_scurve = 24
    ws_sched.merge_cells(f"B{r_mo_scurve}:C{r_mo_scurve}")
    ws_sched[f"B{r_mo_scurve}"] = "CUMULATIVE S-CURVE PROGRESS (%)"
    ws_sched[f"B{r_mo_scurve}"].font = font_data_bold
    ws_sched[f"B{r_mo_scurve}"].alignment = align_right

    ws_sched.cell(row=r_mo_scurve, column=4, value=1.0).alignment = align_right
    ws_sched.cell(row=r_mo_scurve, column=4).font = font_data_bold
    ws_sched.cell(row=r_mo_scurve, column=4).number_format = FMT_PERCENT

    for m in range(1, 25):
        col_m = 4 + m
        col_letter = get_column_letter(col_m)
        cell = ws_sched.cell(row=r_mo_scurve, column=col_m, value=f"={col_letter}{r_mo_cum}/$AB${r_mo_cum}")
        cell.alignment = align_right
        cell.font = font_data_bold
        cell.number_format = FMT_PERCENT
        cell.border = border_cell

    for c in [2, 3, 4]:
        ws_sched.cell(row=r_mo_scurve, column=c).border = border_cell
    ws_sched.row_dimensions[r_mo_scurve].height = 22

    # Column widths for Schedule
    ws_sched.column_dimensions["A"].width = 3
    ws_sched.column_dimensions["B"].width = 8
    ws_sched.column_dimensions["C"].width = 34
    ws_sched.column_dimensions["D"].width = 22
    for m in range(1, 25):
        col_l = get_column_letter(4 + m)
        ws_sched.column_dimensions[col_l].width = 16

    # =========================================================================
    # SHEET 6: PROJECT MAP & SCOPE (EMBEDDED IMAGE & SPECIFICATIONS)
    # =========================================================================
    ws_map = wb.create_sheet(title="Project Map & Scope")
    ws_map.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_map.merge_cells("B2:H2")
    ws_map["B2"] = "CAPM TAU MINE #10 - SITE MAP & DETAILED CONTRACT SCOPE"
    ws_map["B2"].font = font_main_title
    ws_map["B2"].fill = fill_main_title
    ws_map["B2"].alignment = align_title
    ws_map.row_dimensions[2].height = 32

    ws_map.merge_cells("B3:H3")
    ws_map["B3"] = "Total Plant Footprint: 1,463,268.00 m² | 8 Surface Phases + Shaft Filling & Capping | Depth: ±1.5m below NGL"
    ws_map["B3"].font = font_sub_title
    ws_map["B3"].fill = fill_sec_hdr
    ws_map["B3"].alignment = align_center
    ws_map.row_dimensions[3].height = 20

    # Embed Extracted Site Map Image
    img_path = "capm_tau_sitemap_0.png"
    if os.path.exists(img_path):
        try:
            img = OpenpyxlImage(img_path)
            img.width = 750
            img.height = 382
            ws_map.add_image(img, "B5")
        except Exception as e:
            print(f"Image insert notice: {e}")

    # Scope Table below image
    r_map_scope = 26
    ws_map.merge_cells(f"B{r_map_scope}:H{r_map_scope}")
    ws_map[f"B{r_map_scope}"] = "CAPM TAU MINE #10 - DETAILED CONTRACTUAL SCOPE OF WORK BY WORK PACKAGE"
    ws_map[f"B{r_map_scope}"].font = font_sec_hdr
    ws_map[f"B{r_map_scope}"].fill = fill_sec_hdr
    ws_map[f"B{r_map_scope}"].alignment = align_left
    ws_map.row_dimensions[r_map_scope].height = 24

    map_scope_hdrs = ["Phase Ref", "Surface Area (m²)", "Contract Duration", "Detailed Contractual Demolition, Filling & Rehabilitation Scope Specifications"]
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
    ws_map.cell(row=r_map_scope+1, column=5, value=map_scope_hdrs[3]).font = font_tbl_hdr
    ws_map.cell(row=r_map_scope+1, column=5).fill = fill_main_title
    ws_map.cell(row=r_map_scope+1, column=5).alignment = align_left
    for c in range(5, 9):
        ws_map.cell(row=r_map_scope+1, column=c).border = border_cell
    ws_map.row_dimensions[r_map_scope+1].height = 25

    map_scope_data = [
        ("Phase 01", 178860.00, "+-3 Months", "Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Break up all concrete slabs and clear salvage yard; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all material and stockpile at designated area; Shaping / Levelling and cleaning off area"),
        ("Phase 02", 102294.00, "+-4 Months", "Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Breakup all main office buildings and dismantling of workshops steel structures and cooling towers; Breaking of water reservoir for cooling towers; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all material and stockpile at designated area; Shaping / Levelling and cleaning off area"),
        ("Phase 03", 141802.00, "+-4 Months", "Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Breakup main kitchen and surrounding office buildings; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all material and stockpile at designated area; Shaping / Levelling and cleaning off area"),
        ("Phase 04", 12763.00, "+-4 Months", "Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Breakup hostel area; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all material and stockpile at designated area; Shaping / Levelling and cleaning off area"),
        ("Phase 05", 119515.00, "+-3 Months", "Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Breakup and dismantling of workshops steel structures and water reservoirs, and cleanup the salvage yard; Breakup electrical station yard; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all material and stockpile at designated area; Shaping / Levelling and cleaning off area"),
        ("Phase 06", 91628.00, "+-6 Months", "Dismantling of conveyor structure steel structures and breaking concrete columns; Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Breakup two silos' structures; Dismantling steel structures for winder houses and breakup of concrete structures; Breakup all concrete slab on the ground around the shaft area; Demolition of two head gear structures; Demolition of vent shaft structure; Breakup refrigeration plant; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all material and stockpile at designated area; Shaping / Levelling and cleaning off area"),
        ("Phase 07", 264040.00, "+-2 Months", "Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Breakup dams and water reservoir; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all material and stockpile at designated area; Shaping / Levelling and cleaning off area"),
        ("Phase 08", 552366.00, "+-12 Months", "Breakup & Stockpile All Concrete and brickwork above and ±1.5 meters below natural ground level; Waste rock dump must be processed; Cut up and remove all reinforcing material / steel structures; Remove / Load / Haul all material and stockpile at designated area; Shaping / Levelling and cleaning off area"),
        ("Specialist", 0.0, "+-6 Months", "Rehabilitation of CAPM Tau Mine #10 shaft filling (Main Shaft and RV Shaft): Transport all material next to two shafts and stockpile for purpose of shaft filling (10m diameter x 1774m deep = 557,036 m³ x 1.7 = 946,961.2 tons); Capping of the main shafts with reinforced concrete (12m x 12m x 2m = 288 m³); Transport all material next to vent shaft and stockpile for purpose of shaft filling (14m x 14m x 1774m x 1.7 = 591,096.8 tons); Capping of RV Shafts with reinforced steel concrete (14m x 14m x 2m = 392 m³)")
    ]

    for idx, row_data in enumerate(map_scope_data, start=r_map_scope+2):
        pref, parea, pdur, pscope = row_data
        ws_map.cell(row=idx, column=2, value=pref).alignment = align_center
        ws_map.cell(row=idx, column=2).font = font_data_bold
        ws_map.cell(row=idx, column=2).border = border_cell

        c_a = ws_map.cell(row=idx, column=3, value=parea)
        c_a.alignment = align_right
        c_a.font = font_data
        c_a.number_format = FMT_QTY_DEC if parea > 0 else '"N/A"'
        c_a.border = border_cell

        ws_map.cell(row=idx, column=4, value=pdur).alignment = align_center
        ws_map.cell(row=idx, column=4).font = font_data
        ws_map.cell(row=idx, column=4).border = border_cell

        ws_map.merge_cells(f"E{idx}:H{idx}")
        ws_map.cell(row=idx, column=5, value=pscope).alignment = align_left_wrap
        ws_map.cell(row=idx, column=5).font = font_data
        for c in range(5, 9):
            ws_map.cell(row=idx, column=c).border = border_cell

        fill_r = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)
        if fill_r.fill_type:
            for c in range(2, 9):
                ws_map.cell(row=idx, column=c).fill = fill_r
        ws_map.row_dimensions[idx].height = 42

    # Map Total Area Row
    r_map_tot = r_map_scope + 2 + len(map_scope_data)
    ws_map.cell(row=r_map_tot, column=2, value="TOTAL MINE AREA").alignment = align_center
    ws_map.cell(row=r_map_tot, column=2).font = font_grandtotal

    c_map_tot_val = ws_map.cell(row=r_map_tot, column=3, value=f"=SUM(C{r_map_scope+2}:C{r_map_tot-1})")
    c_map_tot_val.alignment = align_right
    c_map_tot_val.font = font_grandtotal
    c_map_tot_val.number_format = FMT_QTY_DEC

    ws_map.cell(row=r_map_tot, column=4, value="24 Mos Integrated").alignment = align_center
    ws_map.cell(row=r_map_tot, column=4).font = font_grandtotal

    ws_map.merge_cells(f"E{r_map_tot}:H{r_map_tot}")
    ws_map.cell(row=r_map_tot, column=5, value="Reconciles exactly to contract document: 1,463,268.00 m² total surface area + 1,538,058.0 tons deep shaft filling.").alignment = align_left
    ws_map.cell(row=r_map_tot, column=5).font = font_grandtotal

    for c in range(2, 9):
        ws_map.cell(row=r_map_tot, column=c).fill = fill_grandtotal
        ws_map.cell(row=r_map_tot, column=c).border = border_grandtotal
    ws_map.row_dimensions[r_map_tot].height = 26

    # Column widths for Map & Scope
    ws_map.column_dimensions["A"].width = 3
    ws_map.column_dimensions["B"].width = 14
    ws_map.column_dimensions["C"].width = 20
    ws_map.column_dimensions["D"].width = 18
    ws_map.column_dimensions["E"].width = 24
    ws_map.column_dimensions["F"].width = 24
    ws_map.column_dimensions["G"].width = 24
    ws_map.column_dimensions["H"].width = 24

    # Save Workbook
    wb.save(output_filename)
    print(f"Workbook successfully saved to: {output_filename}")

if __name__ == "__main__":
    build_complete_capm_tau_mine_model()
