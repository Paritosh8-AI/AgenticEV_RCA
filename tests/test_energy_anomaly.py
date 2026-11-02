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
    assert any(a["type"] == "SPIKE_POWER_EXCEEDED" for a in analysis["anomalies"])
    assert "Modbus" in rca["primary_cause"] or "Surge" in rca["primary_cause"]


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
        meter_records=records
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
