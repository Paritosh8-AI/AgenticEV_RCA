import json
import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

def build_excel_report(json_path="scratch/ocpi_deep_dive_metrics_30d.json", output_path="ElectreeFi_OCPI_Roaming_Metrics_Analysis.xlsx"):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active

    # Styling definitions
    font_title = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    font_section = Font(name="Calibri", size=13, bold=True, color="1B365D")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_sub_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True, color="1E293B")
    font_regular = Font(name="Calibri", size=10, color="1E293B")
    font_kpi_label = Font(name="Calibri", size=10, bold=True, color="64748B")
    font_kpi_val = Font(name="Calibri", size=16, bold=True, color="1B365D")
    font_danger = Font(name="Calibri", size=10, bold=True, color="991B1B")

    fill_navy = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_slate = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_card = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_alt = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_danger_light = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_warning_light = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_success_light = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    card_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # =========================================================================
    # SHEET 1: Executive KPI & Party Metrics
    # =========================================================================
    ws1 = wb.create_sheet(title="Executive & Party Metrics")
    ws1.views.sheetView[0].showGridLines = True

    # Title Banner
    ws1.merge_cells("A1:K2")
    t_cell = ws1["A1"]
    t_cell.value = "ELECTREEFI CMS — OCPI ROAMING PORTAL DEEP DIVE & SUCCESS RATE ANALYSIS"
    t_cell.font = font_title
    t_cell.fill = fill_navy
    t_cell.alignment = align_center

    ws1.merge_cells("A3:K3")
    sub_cell = ws1["A3"]
    sub_cell.value = f"Data Scope: Roaming / OCPI Reservations ({data['window']}) | Trailing 30-Day Complete Audit ({data['total_bookings']:,} Bookings)"
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="64748B")
    sub_cell.alignment = align_left

    # KPI Banner (Row 5 - 6)
    kpis = [
        ("Total Roaming Bookings", f"{data['total_bookings']:,}", "100.0%", "B5", "B6"),
        ("Successful Charges (>=1 kWh)", f"{data['completed_healthy']:,}", f"{data['completed_healthy']/data['total_bookings']*100:.1f}%", "D5", "D6"),
        ("Total Failed / Compromised", f"{data['total_failures']:,}", f"{data['total_failures']/data['total_bookings']*100:.1f}%", "F5", "F6"),
        ("Cancelled Reservations", f"{data['total_cancelled']:,}", f"{data['total_cancelled']/data['total_bookings']*100:.1f}%", "H5", "H6"),
        ("Pre-Charge Aborts (0.0 kWh)", f"{data['completed_zero']:,}", f"{data['completed_zero']/data['total_bookings']*100:.1f}%", "J5", "J6"),
    ]

    for label, val, sub, top_l, bot_l in kpis:
        c1 = ws1[top_l]
        c1.value = label
        c1.font = font_kpi_label
        c1.alignment = align_center
        c1.fill = fill_card

        c2 = ws1[bot_l]
        c2.value = f"{val}  ({sub})"
        c2.font = font_kpi_val
        c2.alignment = align_center
        c2.fill = fill_card

    # Table 1: Party ID Ranking Table
    ws1.cell(row=8, column=1, value="1. ROAMING PARTNER (PARTY ID) COMPREHENSIVE PERFORMANCE & FAILURE AUDIT").font = font_section

    headers1 = [
        "Party ID", "Partner / Gateway Name", "Total Bookings",
        "Successful Sessions (>=1 kWh)", "Success Rate (%)",
        "Total Failed Sessions", "Failure Rate (%)",
        "Cancelled Sessions", "Cancellation Rate (%)",
        "Pre-Charge / Low Aborts", "Abort Rate (%)",
        "Primary Failure Root Cause Issue"
    ]

    for c_idx, h_text in enumerate(headers1, 1):
        cell = ws1.cell(row=10, column=c_idx, value=h_text)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = thin_border
    ws1.row_dimensions[10].height = 25

    row_num = 11
    parties_sorted = sorted(data["party_stats"].items(), key=lambda x: x[1]["total"], reverse=True)
    for pid, p in parties_sorted:
        tot = p["total"]
        succ = p["healthy"]
        fail = p["cancelled"] + p["low"]
        canc = p["cancelled"]
        low = p["low"]
        succ_rate = p["success_rate"] / 100.0
        fail_rate = p["failure_rate"] / 100.0
        canc_rate = p["cancellation_rate"] / 100.0
        abort_rate = p["abort_rate"] / 100.0
        top_reason = p["top_reasons"][0][0] if p["top_reasons"] else "N/A"

        row_fill = fill_alt if row_num % 2 == 0 else fill_white
        # Highlight critical failure partners (>60% failure with >=5 bookings)
        if tot >= 5 and p["failure_rate"] >= 70.0:
            row_fill = fill_danger_light
        elif tot >= 5 and p["failure_rate"] >= 50.0:
            row_fill = fill_warning_light

        values = [
            (pid, align_center, font_bold),
            (p["name"], align_left, font_regular),
            (tot, align_right, font_bold),
            (succ, align_right, font_regular),
            (succ_rate, align_right, font_bold),
            (fail, align_right, font_bold),
            (fail_rate, align_right, font_bold),
            (canc, align_right, font_regular),
            (canc_rate, align_right, font_regular),
            (low, align_right, font_regular),
            (abort_rate, align_right, font_regular),
            (top_reason, align_left, font_regular)
        ]

        for c_idx, (val, al, fnt) in enumerate(values, 1):
            cell = ws1.cell(row=row_num, column=c_idx, value=val)
            cell.font = fnt
            cell.alignment = al
            cell.fill = row_fill
            cell.border = thin_border
            if c_idx in [5, 7, 9, 11]:
                cell.number_format = "0.0%"
            elif c_idx in [3, 4, 6, 8, 10]:
                cell.number_format = "#,##0"

        ws1.row_dimensions[row_num].height = 20
        row_num += 1

    # Total Summary Row for Parties
    ws1.cell(row=row_num, column=1, value="TOTAL / AVERAGE").font = font_header
    ws1.cell(row=row_num, column=2, value="All 34 Roaming Partners").font = font_header
    ws1.cell(row=row_num, column=3, value=data['total_bookings']).font = font_header
    ws1.cell(row=row_num, column=4, value=data['completed_healthy']).font = font_header
    ws1.cell(row=row_num, column=5, value=data['completed_healthy']/data['total_bookings']).font = font_header
    ws1.cell(row=row_num, column=6, value=data['total_failures']).font = font_header
    ws1.cell(row=row_num, column=7, value=data['total_failures']/data['total_bookings']).font = font_header
    ws1.cell(row=row_num, column=8, value=data['total_cancelled']).font = font_header
    ws1.cell(row=row_num, column=9, value=data['total_cancelled']/data['total_bookings']).font = font_header
    ws1.cell(row=row_num, column=10, value=data['completed_low']).font = font_header
    ws1.cell(row=row_num, column=11, value=data['completed_low']/data['total_bookings']).font = font_header
    ws1.cell(row=row_num, column=12, value="Network Average Performance").font = font_header

    for c in range(1, 13):
        cell = ws1.cell(row=row_num, column=c)
        cell.fill = fill_slate
        cell.border = thin_border
        if c in [3, 4, 6, 8, 10]:
            cell.alignment = align_right
            cell.number_format = "#,##0"
        elif c in [5, 7, 9, 11]:
            cell.alignment = align_right
            cell.number_format = "0.0%"
        else:
            cell.alignment = align_left
    ws1.row_dimensions[row_num].height = 22

    # =========================================================================
    # SHEET 2: Station Performance
    # =========================================================================
    ws2 = wb.create_sheet(title="Station Performance")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:K2")
    t2 = ws2["A1"]
    t2.value = "OCPI ROAMING — STATION-LEVEL SUCCESS & FAILURE PERFORMANCE RANKING"
    t2.font = font_title
    t2.fill = fill_navy
    t2.alignment = align_center

    headers2 = [
        "Station Name", "EVSEs", "Total Bookings",
        "Successful Sessions", "Success Rate (%)",
        "Total Failed Sessions", "Failure Rate (%)",
        "Cancelled Sessions", "Pre-Charge Aborts (<1kWh)",
        "Zero-Energy Aborts (0.0 kWh)", "Dominant Roaming Partner",
        "Primary Root Cause Issue"
    ]

    for c_idx, h_text in enumerate(headers2, 1):
        cell = ws2.cell(row=4, column=c_idx, value=h_text)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = thin_border
    ws2.row_dimensions[4].height = 25

    row_num = 5
    stations_sorted = sorted(data["station_stats"].items(), key=lambda x: (x[1]["cancelled"] + x[1]["low"]), reverse=True)
    for st, s in stations_sorted:
        tot = s["total"]
        succ = s["healthy"]
        fail = s["cancelled"] + s["low"]
        canc = s["cancelled"]
        low = s["low"]
        zero = s["zero"]
        succ_rate = s["success_rate"] / 100.0
        fail_rate = s["failure_rate"] / 100.0
        top_p = f"{s['top_parties'][0][0]} ({s['top_parties'][0][1]})" if s['top_parties'] else "N/A"
        top_r = s["top_reasons"][0][0] if s["top_reasons"] else "N/A"

        row_fill = fill_alt if row_num % 2 == 0 else fill_white
        if tot >= 5 and s["failure_rate"] >= 70.0:
            row_fill = fill_danger_light
        elif tot >= 5 and s["failure_rate"] >= 50.0:
            row_fill = fill_warning_light

        values = [
            (st, align_left, font_regular),
            (s["chargers_count"], align_center, font_regular),
            (tot, align_right, font_bold),
            (succ, align_right, font_regular),
            (succ_rate, align_right, font_bold),
            (fail, align_right, font_bold),
            (fail_rate, align_right, font_bold),
            (canc, align_right, font_regular),
            (low, align_right, font_regular),
            (zero, align_right, font_regular),
            (top_p, align_left, font_regular),
            (top_r, align_left, font_regular)
        ]

        for c_idx, (val, al, fnt) in enumerate(values, 1):
            cell = ws2.cell(row=row_num, column=c_idx, value=val)
            cell.font = fnt
            cell.alignment = al
            cell.fill = row_fill
            cell.border = thin_border
            if c_idx in [5, 7]:
                cell.number_format = "0.0%"
            elif c_idx in [2, 3, 4, 6, 8, 9, 10]:
                cell.number_format = "#,##0"

        ws2.row_dimensions[row_num].height = 20
        row_num += 1

    # =========================================================================
    # SHEET 3: Charger Performance
    # =========================================================================
    ws3 = wb.create_sheet(title="Charger Performance")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:K2")
    t3 = ws3["A1"]
    t3.value = "OCPI ROAMING — CHARGER (EVSE)-LEVEL SUCCESS & FAILURE AUDIT"
    t3.font = font_title
    t3.fill = fill_navy
    t3.alignment = align_center

    headers3 = [
        "Charger ID", "Station Name", "Total Bookings",
        "Successful Sessions", "Success Rate (%)",
        "Total Failed Sessions", "Failure Rate (%)",
        "Cancelled Sessions", "Pre-Charge Aborts (<1kWh)",
        "Zero-Energy Aborts (0.0 kWh)", "Dominant Partner",
        "Primary Root Cause Issue"
    ]

    for c_idx, h_text in enumerate(headers3, 1):
        cell = ws3.cell(row=4, column=c_idx, value=h_text)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = thin_border
    ws3.row_dimensions[4].height = 25

    row_num = 5
    chargers_sorted = sorted(data["charger_stats"].items(), key=lambda x: (x[1]["cancelled"] + x[1]["low"]), reverse=True)
    for ch, c in chargers_sorted:
        tot = c["total"]
        succ = c["healthy"]
        fail = c["cancelled"] + c["low"]
        canc = c["cancelled"]
        low = c["low"]
        zero = c["zero"]
        succ_rate = c["success_rate"] / 100.0
        fail_rate = c["failure_rate"] / 100.0
        top_p = f"{c['top_parties'][0][0]} ({c['top_parties'][0][1]})" if c['top_parties'] else "N/A"
        top_r = c["top_reasons"][0][0] if c["top_reasons"] else "N/A"

        row_fill = fill_alt if row_num % 2 == 0 else fill_white
        if tot >= 5 and c["failure_rate"] >= 70.0:
            row_fill = fill_danger_light
        elif tot >= 5 and c["failure_rate"] >= 50.0:
            row_fill = fill_warning_light

        values = [
            (ch, align_left, font_bold),
            (c["station"], align_left, font_regular),
            (tot, align_right, font_bold),
            (succ, align_right, font_regular),
            (succ_rate, align_right, font_bold),
            (fail, align_right, font_bold),
            (fail_rate, align_right, font_bold),
            (canc, align_right, font_regular),
            (low, align_right, font_regular),
            (zero, align_right, font_regular),
            (top_p, align_left, font_regular),
            (top_r, align_left, font_regular)
        ]

        for c_idx, (val, al, fnt) in enumerate(values, 1):
            cell = ws3.cell(row=row_num, column=c_idx, value=val)
            cell.font = fnt
            cell.alignment = al
            cell.fill = row_fill
            cell.border = thin_border
            if c_idx in [5, 7]:
                cell.number_format = "0.0%"
            elif c_idx in [3, 4, 6, 8, 9, 10]:
                cell.number_format = "#,##0"

        ws3.row_dimensions[row_num].height = 20
        row_num += 1

    # =========================================================================
    # SHEET 4: Failure Root Cause Taxonomy & Attribution
    # =========================================================================
    ws4 = wb.create_sheet(title="Failure Taxonomy & Attribution")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:G2")
    t4 = ws4["A1"]
    t4.value = "OCPI ROAMING — ROOT CAUSE TAXONOMY, OWNERSHIP & MITIGATION ACTIONS"
    t4.font = font_title
    t4.fill = fill_navy
    t4.alignment = align_center

    headers4 = [
        "Failure / RCA Category", "Incident Count", "% of All Bookings",
        "% of Total Failures", "Primary Ownership", "Technical Trigger Mechanism",
        "Recommended Engineering & Operational Action"
    ]

    for c_idx, h_text in enumerate(headers4, 1):
        cell = ws4.cell(row=4, column=c_idx, value=h_text)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = thin_border
    ws4.row_dimensions[4].height = 25

    rca_meta = {
        "Zero Energy Delivered (0.0 kWh) - Pre-Charge Abort": {
            "ownership": "CPO / Hardware / Gateway",
            "trigger": "Physical cable coupled & OCPI session registered, but session terminated during parameter handshake, EVSE safety isolation test, or pre-charge voltage ramp.",
            "action": "Audit OCPP StopTransaction meter values; check PLC/CP PWM lock timings with HPCL (Kazam) and IOCL chargers; increase handshake timeout from 30s to 60s."
        },
        "Cancelled by Scheduler (Reservation Expired / Driver No-Show)": {
            "ownership": "Driver / eMSP App Navigation",
            "trigger": "Reservation holding window (15-30 mins) elapsed without the vehicle arriving or initiating an authorized OCPI session on the designated connector.",
            "action": "Implement proactive push notifications at T-10m and T-2m; provide real-time traffic delay extension button in eMSP apps (VinFast, ChargeMod, etc.)."
        },
        "Cancelled by User (eMSP In-App Cancellation)": {
            "ownership": "Driver / eMSP In-App UX",
            "trigger": "Driver proactively pressed 'Cancel' in their roaming consumer app prior to arrival, typically due to finding another charger or route changes.",
            "action": "Enforce cancellation reason prompts in eMSP app; introduce anti-ghost booking deposits or limit concurrent unverified reservations."
        },
        "Cancelled by BecknUser (ONDC/Beckn In-App Cancellation)": {
            "ownership": "Beckn / ONDC Buyer App",
            "trigger": "Buyer cancelled reservation from an ONDC/Beckn protocol consumer application prior to session initiation.",
            "action": "Analyze ONDC buyer interface timeouts; ensure live connector status synchronization between CMS and Beckn BAP/BPP gateway."
        },
        "Canceled by Invalid Session (CPO / Protocol Rejection)": {
            "ownership": "CPO Gateway / Roaming Protocol",
            "trigger": "Destination CPO rejected StartSession command or returned invalid session state due to token validation failure, EVSE busy state, or schema mismatch.",
            "action": "Conduct immediate bilateral OCPI protocol audit with Chargezone (TCZ), HPCL (HPL), and YoCharge (KRI); verify EVSE status mapping and token authorization caching."
        },
        "Partial Low Transfer (<1 kWh)": {
            "ownership": "Vehicle BMS / Connector Coupling",
            "trigger": "Session started and transferred minor energy (<1 kWh) before premature vehicle uncoupling, emergency stop press, or BMS abort.",
            "action": "Inspect connector mechanical locking pins; review OEM BMS error codes (especially Tata Nexon EV and Tigor EV) for early disconnection flags."
        }
    }

    row_num = 5
    total_fails = data["total_failures"]
    total_books = data["total_bookings"]
    for r, cnt in sorted(data["all_reasons"].items(), key=lambda x: x[1], reverse=True):
        meta = rca_meta.get(r, {
            "ownership": "Roaming Gateway / Operator",
            "trigger": "Generic operational or protocol cancellation trigger.",
            "action": "Investigate transaction logs in /Roaming/OCPISession."
        })
        row_fill = fill_alt if row_num % 2 == 0 else fill_white

        ws4.cell(row=row_num, column=1, value=r).font = font_bold
        ws4.cell(row=row_num, column=2, value=cnt).font = font_bold
        ws4.cell(row=row_num, column=3, value=cnt/total_books).font = font_regular
        ws4.cell(row=row_num, column=4, value=cnt/total_fails).font = font_bold
        ws4.cell(row=row_num, column=5, value=meta["ownership"]).font = font_regular
        ws4.cell(row=row_num, column=6, value=meta["trigger"]).font = font_regular
        ws4.cell(row=row_num, column=7, value=meta["action"]).font = font_regular

        ws4.cell(row=row_num, column=1).alignment = align_left
        ws4.cell(row=row_num, column=2).alignment = align_right
        ws4.cell(row=row_num, column=3).alignment = align_right
        ws4.cell(row=row_num, column=4).alignment = align_right
        ws4.cell(row=row_num, column=5).alignment = align_center
        ws4.cell(row=row_num, column=6).alignment = align_left
        ws4.cell(row=row_num, column=7).alignment = align_left

        ws4.cell(row=row_num, column=2).number_format = "#,##0"
        ws4.cell(row=row_num, column=3).number_format = "0.0%"
        ws4.cell(row=row_num, column=4).number_format = "0.0%"

        for c in range(1, 8):
            ws4.cell(row=row_num, column=c).fill = row_fill
            ws4.cell(row=row_num, column=c).border = thin_border

        ws4.row_dimensions[row_num].height = 28
        row_num += 1

    # =========================================================================
    # Auto-fit column widths across all sheets
    # =========================================================================
    for sheet in [ws1, ws2, ws3, ws4]:
        for col in sheet.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                val = str(cell.value or "")
                # Ignore merged cells in row 1-3
                if cell.row in [1, 2, 3]:
                    continue
                max_len = max(max_len, len(val))
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Specific custom column overrides
    ws1.column_dimensions["A"].width = 12
    ws1.column_dimensions["B"].width = 36
    ws1.column_dimensions["L"].width = 45

    ws2.column_dimensions["A"].width = 45
    ws2.column_dimensions["K"].width = 25
    ws2.column_dimensions["L"].width = 45

    ws3.column_dimensions["A"].width = 25
    ws3.column_dimensions["B"].width = 40
    ws3.column_dimensions["K"].width = 25
    ws3.column_dimensions["L"].width = 45

    ws4.column_dimensions["A"].width = 38
    ws4.column_dimensions["E"].width = 25
    ws4.column_dimensions["F"].width = 50
    ws4.column_dimensions["G"].width = 50

    # Save
    if default_sheet.title == "Sheet":
        wb.remove(default_sheet)
    try:
        wb.save(output_path)
        print(f"SUCCESS: Created professional multi-sheet Excel report: {output_path}")
    except PermissionError:
        alt_path = "ElectreeFi_OCPI_Roaming_Metrics_Analysis_30Days.xlsx"
        wb.save(alt_path)
        print(f"NOTE: {output_path} is currently open in Excel. Saved updated 30-day report to: {alt_path}")

if __name__ == "__main__":
    build_excel_report()
