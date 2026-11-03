"""
Unit tests for Energy Anomaly & Power Spike Detector.
Validates:
- Vehicle battery capacity and charging limit resolution
- Charger hardware rating and connector limits
- Consecutive interval meter calculations (Delta E, Delta t, kW)
- Anomaly detectors: Overpower spike, Battery overrun, Meter rollback, Freeze
- RCA diagnostic reporting and financial impact calculation
"""

import pytest
from datetime import datetime, timedelta
from src.rca.energy_anomaly_detector import (
    resolve_vehicle_battery,
    resolve_charger_specs,
    analyze_transaction_energy,
    KNOWN_VEHICLE_BATTERIES
)


def test_vehicle_battery_resolution():
    # 1. Known exact models
    tiago = resolve_vehicle_battery("Tata Tiago EV")
    assert tiago["capacity_kwh"] == 24.0
    assert tiago["max_dc_kw"] == 25.0

    nexon_max = resolve_vehicle_battery("Tata Nexon EV Max")
    assert nexon_max["capacity_kwh"] == 40.5
    assert nexon_max["max_dc_kw"] == 50.0

    curvv = resolve_vehicle_battery("Tata Curvv EV")
    assert curvv["capacity_kwh"] == 55.0

    mg_zs = resolve_vehicle_battery("MG ZS EV")
    assert mg_zs["capacity_kwh"] == 50.3

    byd_atto = resolve_vehicle_battery("BYD Atto 3")
    assert byd_atto["capacity_kwh"] == 60.5

    xev_9e = resolve_vehicle_battery("Mahindra XEV 9e")
    assert xev_9e["capacity_kwh"] == 79.0

    # 2. Unknown fallback
    unknown = resolve_vehicle_battery("Random Concept EV 2030")
    assert unknown["capacity_kwh"] == 50.0


def test_charger_specs_resolution():
    # IOCL Zetwerk 120kW charger
    ioc_specs = resolve_charger_specs("IOC", "IOCLPBDC7784", "9085")
    assert ioc_specs["rated_capacity_kw"] == 120.0
    assert ioc_specs["gun_capacity_kw"] == 60.0
    assert "Zetwerk" in ioc_specs["manufacturer"]

    # Fallback heuristic resolution
    dc_specs = resolve_charger_specs("IOC", "TESTDC60")
    assert dc_specs["rated_capacity_kw"] == 60.0


def test_nominal_charging_session():
    # 20 intervals of clean 60 kW charging
    base_time = datetime(2026, 11, 2, 10, 0, 0)
    val_wh = 5000000  # 5000 kWh start
    records = []

    for i in range(20):
        t_str = (base_time + timedelta(seconds=i * 15)).strftime("%Y-%m-%dT%H:%M:%S+05:30")
        # 0.25 kWh every 15s = (0.25 * 3600 / 15) = 60 kW
        val_wh += 250000
        records.append({
            "Value_Timestamp": t_str,
            "Value": str(val_wh),
            "EnergyInkWh": val_wh / 1000000.0,
            "Location": "Outlet",
            "Measurand": "Energy.Active.Import.Register"
        })

    analysis = analyze_transaction_energy(
        party_id="IOC",
        transaction_id="10001",
        charger_code="TESTDC60",
        start_time="2026-11-02 10:00:00",
        stop_time="2026-11-02 10:05:00",
        vehicle_model_hint="Tata Nexon EV Max",
        meter_records=records
    )

    summary = analysis["summary"]
    rca = analysis["rca_report"]

    assert summary["is_anomaly"] is False
    assert summary["highest_severity"] == "NOMINAL"
    assert summary["spike_count"] == 0
    assert summary["rollback_count"] == 0
    assert pytest.approx(summary["total_energy_kwh"], 0.1) == 5.0
    assert pytest.approx(summary["peak_power_kw"], 0.5) == 60.0
    assert rca["verdict"].startswith("NOMINAL")
    assert rca["financial_impact"]["phantom_kwh"] == 0.0


def test_overpower_spike_detection():
    # Session with a sudden Modbus CT surge (e.g. 720 kW on a 60 kW charger)
    base_time = datetime(2026, 11, 2, 10, 0, 0)
    val_wh = 1000000
    records = []

    for i in range(10):
        t_str = (base_time + timedelta(seconds=i * 15)).strftime("%Y-%m-%dT%H:%M:%S+05:30")
        if i == 4:
            # Jump by 3.0 kWh in 15 seconds -> 720 kW!
            val_wh += 3000000
        else:
            val_wh += 200000
        records.append({
            "Value_Timestamp": t_str,
            "Value": str(val_wh),
            "EnergyInkWh": val_wh / 1000000.0,
            "Location": "Outlet",
            "Measurand": "Energy.Active.Import.Register"
        })

    analysis = analyze_transaction_energy(
        party_id="IOC",
        transaction_id="10002",
        charger_code="TESTDC60",
        start_time="2026-11-02 10:00:00",
        stop_time="2026-11-02 10:02:30",
        vehicle_model_hint="Tata Tiago EV",
        meter_records=records
    )

    summary = analysis["summary"]
    rca = analysis["rca_report"]

    assert summary["is_anomaly"] is True
    assert summary["highest_severity"] == "CRITICAL"
    assert summary["spike_count"] >= 1
    assert any(a["type"] in ("ABNORMAL_ENERGY_JUMP", "SPIKE_POWER_EXCEEDED") for a in analysis["anomalies"])
    assert "Modbus" in rca["primary_cause"] or "Surge" in rca["primary_cause"] or "Erratic" in rca["primary_cause"]


def test_battery_capacity_overrun():
    # Session where 55 kWh is billed into a 24 kWh Tata Tiago EV
    records = [
        {"Value_Timestamp": "2026-11-02T10:00:00+05:30", "Value": "10000000", "EnergyInkWh": 10.0, "Location": "Outlet"},
        {"Value_Timestamp": "2026-11-02T10:45:00+05:30", "Value": "65000000", "EnergyInkWh": 65.0, "Location": "Outlet"}
    ]

    analysis = analyze_transaction_energy(
        party_id="IOC",
        transaction_id="10003",
        charger_code="TESTDC60",
        start_time="2026-11-02 10:00:00",
        stop_time="2026-11-02 10:45:00",
        vehicle_model_hint="Tata Tiago EV",
        meter_records=records,
        tariff_per_unit=14.0
    )

    summary = analysis["summary"]
    rca = analysis["rca_report"]

    assert summary["is_anomaly"] is True
    assert any(a["type"] == "BATTERY_CAPACITY_OVERRUN" for a in analysis["anomalies"])
    assert rca["financial_impact"]["phantom_kwh"] > 0
    assert rca["financial_impact"]["disputed_amount_inr"] > 0


def test_counter_rollback():
    # Session where reading drops (register reset or overflow)
    records = [
        {"Value_Timestamp": "2026-11-02T10:00:00+05:30", "Value": "2500000", "EnergyInkWh": 2.50, "Location": "Outlet"},
        {"Value_Timestamp": "2026-11-02T10:00:15+05:30", "Value": "2600000", "EnergyInkWh": 2.60, "Location": "Outlet"},
        {"Value_Timestamp": "2026-11-02T10:00:30+05:30", "Value": "2100000", "EnergyInkWh": 2.10, "Location": "Outlet"}  # Drop!
    ]

    analysis = analyze_transaction_energy(
        party_id="IOC",
        transaction_id="10004",
        charger_code="TESTDC60",
        start_time="2026-11-02 10:00:00",
        stop_time="2026-11-02 10:01:00",
        vehicle_model_hint="Tata Nexon EV Max",
        meter_records=records
    )

    assert analysis["summary"]["is_anomaly"] is True
    assert analysis["summary"]["rollback_count"] >= 1
    assert any(a["type"] == "METER_ROLLBACK" for a in analysis["anomalies"])


def test_investigate_booking_energy_empty():
    from src.rca.energy_anomaly_detector import investigate_booking_energy
    res = investigate_booking_energy("IOC", "")
    assert res["success"] is False
    assert "No Booking ID provided" in res["error"]


def test_billing_tax_audit_nominal():
    from src.rca.energy_anomaly_detector import compute_billing_tax_audit
    analysis = {
        "summary": {"total_energy_kwh": 12.33, "is_anomaly": False},
        "rca_report": {"financial_impact": {"phantom_kwh": 0.0}},
        "vehicle_specs": {"capacity_kwh": 79.0}
    }
    booking = {
        "PerUnitPrice": 14.0,
        "ActualChargingAmount": 172.55,
        "ActualTax": 31.06,
        "ActualTotalAmount": 203.61,
        "PaymentRecieved": 451.70,
        "RefundedAmount": 248.09,
        "EnergyConsumed_indecimal": 12.33
    }
    audit = compute_billing_tax_audit(analysis, booking, party_id="IOC")
    assert audit["has_billing_data"] is True
    assert audit["tariff_rate_inr_per_kwh"] == 14.0
    assert audit["tax_rate_pct"] == 18.0
    assert audit["actual_user_paid"]["total_amount_after_taxes_inr"] == 203.61
    assert audit["real_should_charge"]["total_amount_after_taxes_inr"] == 203.61
    assert audit["comparison"]["variance_inr"] == 0.0
    assert audit["comparison"]["status"] == "ACCURATELY_CHARGED"


def test_billing_tax_audit_overcharge():
    from src.rca.energy_anomaly_detector import compute_billing_tax_audit
    # User was billed for 50 kWh on a 24 kWh Tata Tiago EV (26 kWh phantom)
    analysis = {
        "summary": {"total_energy_kwh": 50.0, "is_anomaly": True},
        "rca_report": {"financial_impact": {"phantom_kwh": 26.0}},
        "vehicle_specs": {"capacity_kwh": 24.0}
    }
    booking = {
        "PerUnitPrice": 10.0,
        "ActualChargingAmount": 500.0,
        "ActualTax": 90.0,
        "ActualTotalAmount": 590.0,
        "PaymentRecieved": 590.0,
        "RefundedAmount": 0.0,
        "EnergyConsumed_indecimal": 50.0
    }
    audit = compute_billing_tax_audit(analysis, booking, party_id="IOC")
    assert audit["actual_user_paid"]["total_amount_after_taxes_inr"] == 590.0
    # Real legitimate energy = 24 kWh @ 10 = 240 base + 43.20 (18% tax) = 283.20 total
    assert audit["real_should_charge"]["legitimate_energy_kwh"] == 24.0
    assert audit["real_should_charge"]["base_amount_inr"] == 240.0
    assert audit["real_should_charge"]["tax_amount_inr"] == 43.20
    assert audit["real_should_charge"]["total_amount_after_taxes_inr"] == 283.20
    # Overcharged: 590 - 283.20 = 306.80
    assert audit["comparison"]["variance_inr"] == 306.80
    assert audit["comparison"]["base_discrepancy_inr"] == 260.0
    assert audit["comparison"]["tax_discrepancy_inr"] == 46.80
    assert audit["comparison"]["status"] == "OVERCHARGED"
    assert "OVERCHARGED BY" in audit["comparison"]["status_label"]


def test_zero_energy_delivered_anomaly():
    from src.rca.energy_anomaly_detector import analyze_transaction_energy
    # 1 meter reading with 0 kWh delivered
    records = [
        {"Value_Timestamp": "2026-10-11T16:32:20+05:30", "Value": "430590", "EnergyInkWh": 430.59, "Location": "Outlet"}
    ]
    analysis = analyze_transaction_energy(
        party_id="ELC",
        transaction_id="11037320",
        charger_code="EESL-JP-NP-DC-014",
        start_time="2026-10-11 16:32:18",
        stop_time="2026-10-11 16:33:00",
        vehicle_model_hint="Bajaj RE E-TEC 9.0",
        meter_records=records
    )
    assert analysis["summary"]["is_anomaly"] is True
    assert any(a["type"] == "ZERO_ENERGY_DELIVERY" for a in analysis["anomalies"])
    verdict_lower = analysis["rca_report"]["verdict"].lower()
    assert "zero-energy" in verdict_lower or "abort" in verdict_lower


def test_discrepancy_billed_vs_metered_zombie():
    from src.rca.energy_anomaly_detector import analyze_transaction_energy, compute_billing_tax_audit
    # Physical meter had 0 kWh delivered, but CMS records billed 169.34 kWh
    records = [
        {"Value_Timestamp": "2026-10-11T16:32:20+05:30", "Value": "430590", "EnergyInkWh": 430.59, "Location": "Outlet"}
    ]
    tx_meta = {
        "start_value": 430590,
        "stop_value": 599930,
        "TotalMeterReading": 169340.0,
        "EnergyConsumedDouble": 169.34
    }
    booking = {
        "PerUnitPrice": 18.0,
        "ActualChargingAmount": 3048.12,
        "ActualTax": 548.66,
        "ActualTotalAmount": 3596.78,
        "PaymentRecieved": 11.68,
        "PendingAmount": 3585.10,
        "EnergyConsumed_indecimal": 169.34
    }
    analysis = analyze_transaction_energy(
        party_id="ELC",
        transaction_id="11037320",
        charger_code="EESL-JP-NP-DC-014",
        start_time="2026-10-11 16:32:18",
        stop_time="2026-11-01 14:42:00",
        vehicle_model_hint="Bajaj RE E-TEC 9.0",
        meter_records=records,
        tx_meta=tx_meta,
        booking_record=booking,
        tariff_per_unit=18.0
    )
    assert analysis["summary"]["is_anomaly"] is True
    assert any(a["type"] == "DISCREPANCY_BILLED_VS_METERED" for a in analysis["anomalies"])
    assert any(a["type"] == "BATTERY_CAPACITY_OVERRUN" for a in analysis["anomalies"])
    assert any(a["type"] == "ORPHANED_ZOMBIE_TRANSACTION" for a in analysis["anomalies"])

    # Audit check: legitimate energy must be 0.00 kWh and real charge must be 0.00
    audit = compute_billing_tax_audit(analysis, booking, party_id="ELC")
    assert audit["real_should_charge"]["legitimate_energy_kwh"] == 0.0
    assert audit["real_should_charge"]["total_amount_after_taxes_inr"] == 0.0
    assert audit["comparison"]["status"] == "OVERCHARGED"
    assert audit["comparison"]["variance_inr"] == 3596.78
    assert "CRITICAL OVERCHARGE" in audit["comparison"]["status_label"]
