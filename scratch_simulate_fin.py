import openpyxl

rates = {
    'R_EST': 35000.0,
    'R_STL_STD': 21000.0,
    'R_STL_HVY': 28500.0,
    'R_STL_WKP': 19500.0,
    'R_CON_M3': 240.0,
    'R_CON_TON': 100.0,
    'R_HAUL_TON': 120.0,
    'R_SHAPE_M2': 18.5,
    'R_SILO_DEM': 260.0,
    'R_WRD_PROC': 45.0,
    'R_SHAFT_FILL': 55.0,
    'R_SHAFT_CAP': 4200.0,
    'P_PG': 0.10,
    'P_CONT': 0.10,
    'P_PROFIT': 0.05,
    'P_VAT': 0.15,
    'P_DURATION': 24.0,
}

phases = {
    1: {'name': 'Phase 01: Salvage Yard & Slabs', 'items': [
        ('Est', 1 * rates['R_EST']),
        ('Steel', 6 * rates['R_STL_STD']),
        ('Conc', 4500 * rates['R_CON_M3']),
        ('Haul', 10800 * rates['R_HAUL_TON']),
        ('Shape', 178860 * rates['R_SHAPE_M2']),
    ]},
    2: {'name': 'Phase 02: Admin & Cooling Towers', 'items': [
        ('Est', 1 * rates['R_EST']),
        ('Steel', 8 * rates['R_STL_STD']),
        ('Conc', 6200 * rates['R_CON_M3']),
        ('Haul', 14880 * rates['R_HAUL_TON']),
        ('Shape', 102294 * rates['R_SHAPE_M2']),
    ]},
    3: {'name': 'Phase 03: Kitchen & Offices Area', 'items': [
        ('Est', 1 * rates['R_EST']),
        ('Steel', 5 * rates['R_STL_STD']),
        ('Conc', 5800 * rates['R_CON_M3']),
        ('Haul', 13920 * rates['R_HAUL_TON']),
        ('Shape', 141802 * rates['R_SHAPE_M2']),
    ]},
    4: {'name': 'Phase 04: Hostel Complex Area', 'items': [
        ('Est', 1 * rates['R_EST']),
        ('Steel', 6 * rates['R_STL_WKP']),
        ('Conc', 3500 * rates['R_CON_M3']),
        ('Haul', 8400 * rates['R_HAUL_TON']),
        ('Shape', 12763 * rates['R_SHAPE_M2']),
    ]},
    5: {'name': 'Phase 05: Workshops & Substation', 'items': [
        ('Est', 1 * rates['R_EST']),
        ('Steel', 7 * rates['R_STL_WKP']),
        ('Conc', 5200 * rates['R_CON_M3']),
        ('Haul', 12480 * rates['R_HAUL_TON']),
        ('Shape', 119515 * rates['R_SHAPE_M2']),
    ]},
    6: {'name': 'Phase 06: Shafts, Silos & Headgears', 'items': [
        ('Est', 1 * rates['R_EST']),
        ('Steel_HG', 2 * rates['R_STL_HVY']),
        ('Steel_Conv', 4 * rates['R_STL_HVY']),
        ('Steel_Winder', 4 * rates['R_STL_HVY']),
        ('Conc_Silos', 4200 * rates['R_SILO_DEM']),
        ('Conc_Bank', 8600 * rates['R_CON_M3']),
        ('Haul', 30720 * rates['R_HAUL_TON']),
        ('Shape', 91628 * rates['R_SHAPE_M2']),
    ]},
    7: {'name': 'Phase 07: Dams & Water Reservoirs', 'items': [
        ('Est', 1 * rates['R_EST']),
        ('Steel', 4 * rates['R_STL_STD']),
        ('Conc', 7500 * rates['R_CON_M3']),
        ('Haul', 18000 * rates['R_HAUL_TON']),
        ('Shape', 264040 * rates['R_SHAPE_M2']),
    ]},
    8: {'name': 'Phase 08: Waste Rock Dump Processing', 'items': [
        ('Est', 1 * rates['R_EST']),
        ('Steel', 2 * rates['R_STL_STD']),
        ('Conc', 3200 * rates['R_CON_M3']),
        ('WRD_Proc', 1538058 * rates['R_WRD_PROC']),
        ('Haul', 7680 * rates['R_HAUL_TON']),
        ('Shape', 552366 * rates['R_SHAPE_M2']),
    ]},
    9: {'name': 'Specialist: Shaft Filling & Capping', 'items': [
        ('Est', 1 * rates['R_EST']),
        ('Main_Shaft_Fill', 946961.2 * rates['R_SHAFT_FILL']),
        ('Main_Shaft_Cap', 288 * rates['R_SHAFT_CAP']),
        ('RV_Shaft_Fill', 591096.8 * rates['R_SHAFT_FILL']),
        ('RV_Shaft_Cap', 392 * rates['R_SHAFT_CAP']),
    ]}
}

print("=== PHASE DIRECT COSTS ===")
total_direct = 0
for p_idx, p in phases.items():
    p_tot = sum(val for _, val in p['items'])
    total_direct += p_tot
    print(f"{p['name']}: R {p_tot:,.2f}")

print(f"\nTotal Direct Project Spend (Excl P&G): R {total_direct:,.2f}")

pg = total_direct * rates['P_PG']
sub_excl_cont = total_direct + pg
cont = sub_excl_cont * rates['P_CONT']
cost_a = sub_excl_cont + cont
monthly_a = cost_a / rates['P_DURATION']
profit = cost_a * rates['P_PROFIT']
sub_b = cost_a + profit
vat = sub_b * rates['P_VAT']
tender_c = sub_b + vat

print(f"Preliminaries & General (10%): R {pg:,.2f}")
print(f"Subtotal Excl Contingencies: R {sub_excl_cont:,.2f}")
print(f"Contingencies (10%): R {cont:,.2f}")
print(f"Total Project Costs A (Excl. VAT): R {cost_a:,.2f}")
print(f"Monthly Project Cost A Average (24 Mos): R {monthly_a:,.2f}")
print(f"Contractor Profit (5%): R {profit:,.2f}")
print(f"Sub-Total B Tender Price (Excl. VAT): R {sub_b:,.2f}")
print(f"15% SARS VAT: R {vat:,.2f}")
print(f"TOTAL PROJECT COST (TENDER PRICE INCL. VAT): R {tender_c:,.2f}")
