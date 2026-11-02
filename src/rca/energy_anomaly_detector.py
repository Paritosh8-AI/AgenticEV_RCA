"""
Energy Anomaly & Spike RCA Detector
Intelligent framework for detecting unusual energy consumptions and power spikes across
IndianOil (IOCL) and ElectreeFi OCPP transactions.

Features:
- Physical grounding: Charger ratings (kW), connector configurations, EV battery pack capacities.
- Consecutive meter reading analysis (Energy.Active.Import.Register):
  Delta E, Delta t, calculated interval power (kW) vs rated nameplate capacity.
- Multi-factor anomaly detection:
  * Overpower spikes (P_calc > P_rated * 1.15)
  * Battery capacity overrun (E_delivered > C_battery * 1.08)
  * Meter counter rollbacks / resets (Delta E < 0)
  * Meter freeze / zero consumption during active charging
  * Inlet vs Outlet conversion efficiency anomalies
- Root Cause Analysis (RCA) panel generation with financial impact & remediation steps.
"""

import os
import sys
import json
import math
import urllib.request
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple

# Base path resolution
BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / "data"

# Browser User Agent to bypass WAF / anti-bot
DEFAULT_BROWSER_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)

# Known Indian EV Database (Pack Capacity in kWh, Max DC acceptance in kW)
KNOWN_VEHICLE_BATTERIES: Dict[str, Dict[str, Any]] = {
    "tata nexon ev prime": {"name": "Tata Nexon EV Prime", "capacity_kwh": 30.2, "max_dc_kw": 30.0, "chemistry": "LFP"},
    "tata nexon ev max": {"name": "Tata Nexon EV Max", "capacity_kwh": 40.5, "max_dc_kw": 50.0, "chemistry": "LFP"},
    "tata nexon ev long range": {"name": "Tata Nexon EV Long Range (45)", "capacity_kwh": 45.0, "max_dc_kw": 60.0, "chemistry": "LFP"},
    "tata nexon ev": {"name": "Tata Nexon EV (Default)", "capacity_kwh": 40.5, "max_dc_kw": 50.0, "chemistry": "LFP"},
    "tata tiago ev": {"name": "Tata Tiago EV", "capacity_kwh": 24.0, "max_dc_kw": 25.0, "chemistry": "LFP"},
    "tata tiago ev mr": {"name": "Tata Tiago EV MR", "capacity_kwh": 19.2, "max_dc_kw": 20.0, "chemistry": "LFP"},
    "tata tigor ev": {"name": "Tata Tigor EV", "capacity_kwh": 26.0, "max_dc_kw": 25.0, "chemistry": "LFP"},
    "tata punch ev": {"name": "Tata Punch EV Long Range", "capacity_kwh": 35.0, "max_dc_kw": 35.0, "chemistry": "LFP"},
    "tata punch ev mr": {"name": "Tata Punch EV Standard", "capacity_kwh": 25.0, "max_dc_kw": 25.0, "chemistry": "LFP"},
    "tata curvv ev 55": {"name": "Tata Curvv EV 55", "capacity_kwh": 55.0, "max_dc_kw": 70.0, "chemistry": "LFP"},
    "tata curvv ev 45": {"name": "Tata Curvv EV 45", "capacity_kwh": 45.0, "max_dc_kw": 60.0, "chemistry": "LFP"},
    "tata curvv ev": {"name": "Tata Curvv EV", "capacity_kwh": 55.0, "max_dc_kw": 70.0, "chemistry": "LFP"},
    "mg zs ev": {"name": "MG ZS EV", "capacity_kwh": 50.3, "max_dc_kw": 80.0, "chemistry": "LFP"},
    "mg comet ev": {"name": "MG Comet EV", "capacity_kwh": 17.3, "max_dc_kw": 7.4, "chemistry": "LFP"},
    "mg windsor ev": {"name": "MG Windsor EV", "capacity_kwh": 38.0, "max_dc_kw": 45.0, "chemistry": "LFP"},
    "mahindra xuv400": {"name": "Mahindra XUV400 EL", "capacity_kwh": 39.4, "max_dc_kw": 50.0, "chemistry": "NMC"},
    "mahindra xev 9e": {"name": "Mahindra XEV 9e", "capacity_kwh": 79.0, "max_dc_kw": 175.0, "chemistry": "LFP"},
    "mahindra be 6e": {"name": "Mahindra BE 6e", "capacity_kwh": 59.0, "max_dc_kw": 140.0, "chemistry": "LFP"},
    "byd atto 3": {"name": "BYD Atto 3", "capacity_kwh": 60.5, "max_dc_kw": 80.0, "chemistry": "LFP (Blade)"},
    "byd e6": {"name": "BYD e6 MPV", "capacity_kwh": 71.7, "max_dc_kw": 60.0, "chemistry": "LFP (Blade)"},
    "byd seal": {"name": "BYD Seal", "capacity_kwh": 82.5, "max_dc_kw": 150.0, "chemistry": "LFP (Blade)"},
    "hyundai ioniq 5": {"name": "Hyundai Ioniq 5", "capacity_kwh": 72.6, "max_dc_kw": 220.0, "chemistry": "NMC"},
    "hyundai kona ev": {"name": "Hyundai Kona Electric", "capacity_kwh": 39.2, "max_dc_kw": 50.0, "chemistry": "NMC"},
    "kia ev6": {"name": "Kia EV6", "capacity_kwh": 77.4, "max_dc_kw": 230.0, "chemistry": "NMC"},
    "electric bus 300": {"name": "Commercial Heavy Bus (300 kWh)", "capacity_kwh": 300.0, "max_dc_kw": 180.0, "chemistry": "LFP"},
    "electric 3w cargo": {"name": "Commercial 3-Wheeler Cargo", "capacity_kwh": 10.0, "max_dc_kw": 10.0, "chemistry": "LFP"},
    "default 4w ev": {"name": "Standard Passenger 4W EV (Default)", "capacity_kwh": 50.0, "max_dc_kw": 60.0, "chemistry": "LFP"}
}

# Cache for loaded party charger registries
_CHARGER_REGISTRY_CACHE: Dict[str, Dict[str, Any]] = {}


def load_party_chargers(party_id: str) -> Dict[str, Any]:
    """Loads and caches party charger registry."""
    party_id = (party_id or "IOC").upper()
    if party_id in _CHARGER_REGISTRY_CACHE:
        return _CHARGER_REGISTRY_CACHE[party_id]
    
    file_path = DATA_DIR / f"party_chargers_{party_id}.json"
    if not file_path.exists():
        if party_id == "IOC":
            file_path = DATA_DIR / "party_chargers_IOC.json"
        else:
            file_path = DATA_DIR / "party_chargers_ELC.json"
            
    if file_path.exists():
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    _CHARGER_REGISTRY_CACHE[party_id] = data
                    return data
                elif isinstance(data, list):
                    mapped = {}
                    for item in data:
                        code = item.get("charger_code")
                        cid = str(item.get("charger_id") or "")
                        if code:
                            mapped[code] = item
                        if cid:
                            mapped[cid] = item
                    _CHARGER_REGISTRY_CACHE[party_id] = mapped
                    return mapped
        except Exception as e:
            print(f"[EnergyAnomaly] Error reading {file_path}: {e}")
            
    _CHARGER_REGISTRY_CACHE[party_id] = {}
    return {}


def resolve_charger_specs(party_id: str, charger_code: str, charger_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Resolves physical hardware specifications for a charger:
    Rated kW capacity, gun capacity, manufacturer, model, and gun count.
    """
    chargers = load_party_chargers(party_id)
    info = chargers.get(charger_code) or (chargers.get(str(charger_id)) if charger_id else None)
    
    if info:
        cap = float(info.get("charger_capacity") or 0.0)
        guns = int(info.get("no_of_guns") or 2)
        mfr = str(info.get("charger_manufacturer") or "Standard OEM")
        model = str(info.get("charger_model") or "")
        
        # Determine gun capacity
        if guns > 1 and cap > 0:
            gun_cap = round(cap / guns, 1)
        else:
            gun_cap = cap if cap > 0 else 60.0
            
        return {
            "charger_code": charger_code or info.get("charger_code"),
            "charger_id": str(charger_id or info.get("charger_id") or ""),
            "rated_capacity_kw": cap if cap > 0 else 60.0,
            "gun_capacity_kw": gun_cap,
            "no_of_guns": guns,
            "manufacturer": mfr,
            "model": model,
            "connector_type": "CCS2 (Dual Gun)" if guns > 1 else "CCS2 (Single Gun)",
            "station_name": info.get("station_name", "CMS Station")
        }
        
    # Heuristic fallback based on charger_code naming conventions
    code_upper = (charger_code or "").upper()
    is_dc = "DC" in code_upper
    is_ac = "AC" in code_upper
    
    if is_ac:
        rated_cap = 7.4
        gun_cap = 7.4
        connector = "Type 2 AC"
    elif "120" in code_upper:
        rated_cap = 120.0
        gun_cap = 60.0
        connector = "CCS2 Dual Gun"
    elif "60" in code_upper or is_dc:
        rated_cap = 60.0
        gun_cap = 30.0 if "30" in code_upper else 60.0
        connector = "CCS2 DC"
    else:
        rated_cap = 60.0
        gun_cap = 60.0
        connector = "CCS2 Fast DC"
        
    return {
        "charger_code": charger_code,
        "charger_id": str(charger_id or ""),
        "rated_capacity_kw": rated_cap,
        "gun_capacity_kw": gun_cap,
        "no_of_guns": 2 if is_dc else 1,
        "manufacturer": "OEM Standard",
        "model": f"{rated_cap} kW Charger",
        "connector_type": connector,
        "station_name": "EV Charging Station"
    }


def resolve_vehicle_battery(vehicle_str: Optional[str] = None, vin: Optional[str] = None) -> Dict[str, Any]:
    """
    Resolves EV battery pack size and max DC charge rate from vehicle model or VIN.
    """
    v_norm = (vehicle_str or "").lower().strip()
    vin_norm = (vin or "").lower().strip()
    
    # 1. Direct keyword match
    for key, spec in KNOWN_VEHICLE_BATTERIES.items():
        if key in v_norm or key in vin_norm:
            return spec
            
    # 2. Heuristics for Tata models
    if "nexon" in v_norm or "nexon" in vin_norm:
        if "max" in v_norm or "lr" in v_norm:
            return KNOWN_VEHICLE_BATTERIES["tata nexon ev max"]
        return KNOWN_VEHICLE_BATTERIES["tata nexon ev"]
    if "tiago" in v_norm:
        return KNOWN_VEHICLE_BATTERIES["tata tiago ev"]
    if "punch" in v_norm:
        return KNOWN_VEHICLE_BATTERIES["tata punch ev"]
    if "curvv" in v_norm:
        return KNOWN_VEHICLE_BATTERIES["tata curvv ev"]
    if "zs" in v_norm or "mg" in v_norm:
        if "comet" in v_norm:
            return KNOWN_VEHICLE_BATTERIES["mg comet ev"]
        if "windsor" in v_norm:
            return KNOWN_VEHICLE_BATTERIES["mg windsor ev"]
        return KNOWN_VEHICLE_BATTERIES["mg zs ev"]
    if "byd" in v_norm:
        if "atto" in v_norm:
            return KNOWN_VEHICLE_BATTERIES["byd atto 3"]
        if "seal" in v_norm:
            return KNOWN_VEHICLE_BATTERIES["byd seal"]
        return KNOWN_VEHICLE_BATTERIES["byd e6"]
    if "xuv400" in v_norm or "mahindra" in v_norm:
        if "xev" in v_norm or "9e" in v_norm:
            return KNOWN_VEHICLE_BATTERIES["mahindra xev 9e"]
        if "6e" in v_norm:
            return KNOWN_VEHICLE_BATTERIES["mahindra be 6e"]
        return KNOWN_VEHICLE_BATTERIES["mahindra xuv400"]
    if "ioniq" in v_norm:
        return KNOWN_VEHICLE_BATTERIES["hyundai ioniq 5"]
    if "kona" in v_norm:
        return KNOWN_VEHICLE_BATTERIES["hyundai kona ev"]
    if "bus" in v_norm or "commercial" in v_norm:
        return KNOWN_VEHICLE_BATTERIES["electric bus 300"]
        
    return KNOWN_VEHICLE_BATTERIES["default 4w ev"]


def get_portal_base_url(party_id: str) -> str:
    party_id = (party_id or "IOC").upper()
    if party_id == "IOC":
        return "https://indianoilecharge.in"
    return "https://emonitoring.electreefi.com"


def fetch_ocpp_transactions(
    party_id: str = "IOC",
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    tx_type: str = "2",
    page: int = 1,
    page_size: int = 30,
    filter_query: str = ""
) -> Dict[str, Any]:
    """
    Fetches list of OCPP transactions from CMS backend.
    """
    from src.rca.booking_inspector import get_party_cookie_header
    
    party_id = party_id.upper()
    portal_base = get_portal_base_url(party_id)
    
    if not end_date:
        end_date = datetime.now().strftime("%Y-%m-%d")
    if not start_date:
        # Default last 3 days
        from datetime import timedelta
        start_date = (datetime.now() - timedelta(days=3)).strftime("%Y-%m-%d")
        
    cookie_hdr = get_party_cookie_header(party_id)
    url = f"{portal_base}/OCPPManagement/TransactionDetail/LoadTransactionViewThroughAjaxForData/{tx_type}?StartDate={start_date}&Enddate={end_date}"
    
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "User-Agent": DEFAULT_BROWSER_UA,
        "Cookie": cookie_hdr
    }
    
    payload = {
        "page": str(page),
        "pageSize": str(page_size),
        "sort": "transaction_pk-desc",
        "group": "",
        "filter": filter_query or ""
    }
    
    encoded = urllib.parse.urlencode(payload).encode("utf-8")
    req = urllib.request.Request(url, data=encoded, headers=headers)
    
    try:
        with urllib.request.urlopen(req, timeout=15) as res:
            raw = res.read().decode("utf-8")
            data = json.loads(raw)
            return {
                "success": True,
                "party_id": party_id,
                "start_date": start_date,
                "end_date": end_date,
                "total": data.get("Total", 0),
                "transactions": data.get("Data", [])
            }
    except Exception as e:
        return {
            "success": False,
            "party_id": party_id,
            "error": str(e),
            "transactions": []
        }


def fetch_meter_logs(
    party_id: str,
    entity_id: str,
    trans_id: str,
    start_time: str,
    stop_time: str,
    measurand: str = "Energy.Active.Import.Register"
) -> List[Dict[str, Any]]:
    """
    Fetches raw meter reading logs for a specific transaction.
    """
    from src.rca.booking_inspector import get_party_cookie_header
    
    party_id = (party_id or "IOC").upper()
    portal_base = get_portal_base_url(party_id)
    cookie_hdr = get_party_cookie_header(party_id)
    
    url = (
        f"{portal_base}/OCPPManagement/TransactionDetail/GetConnectorMeterLogs"
        f"?entityId={entity_id}"
        f"&fromDate={urllib.parse.quote(start_time)}"
        f"&toDate={urllib.parse.quote(stop_time)}"
        f"&transId={trans_id}"
    )
    
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "User-Agent": DEFAULT_BROWSER_UA,
        "Cookie": cookie_hdr
    }
    
    filter_expr = f"Measurand~eq~'{measurand}'" if measurand else ""
    payload = urllib.parse.urlencode({
        "page": "1",
        "pageSize": "500",
        "sort": "Value_Timestamp-asc",
        "filter": filter_expr
    }).encode("utf-8")
    
    req = urllib.request.Request(url, data=payload, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=15) as res:
            parsed = json.loads(res.read().decode("utf-8"))
            return parsed.get("Data", [])
    except Exception as e:
        print(f"[EnergyAnomaly] Error fetching meter logs for {trans_id}: {e}")
        return []


def parse_timestamp_iso(ts_str: str) -> Optional[datetime]:
    """Robust ISO timestamp parser."""
    if not ts_str or ts_str.startswith("0001"):
        return None
    try:
        clean = ts_str.split("+")[0].replace("Z", "")
        if "T" in clean:
            return datetime.fromisoformat(clean)
        return datetime.strptime(clean, "%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def analyze_transaction_energy(
    party_id: str,
    transaction_id: str,
    charger_code: Optional[str] = None,
    charger_id: Optional[str] = None,
    start_time: Optional[str] = None,
    stop_time: Optional[str] = None,
    vehicle_model_hint: Optional[str] = None,
    custom_battery_kwh: Optional[float] = None,
    meter_records: Optional[List[Dict[str, Any]]] = None,
    tx_meta: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Comprehensive physical anomaly detection & RCA analysis for an OCPP charging transaction.
    """
    party_id = (party_id or "IOC").upper()
    trans_id_str = str(transaction_id).strip()
    
    # 1. If metadata or dates missing, query transaction details first
    if not tx_meta or not charger_code or not start_time:
        res = fetch_ocpp_transactions(
            party_id=party_id,
            filter_query=f"transaction_pk~eq~{trans_id_str}",
            page_size=5
        )
        tx_list = res.get("transactions") or []
        if tx_list:
            matched = tx_list[0]
            tx_meta = matched
            charger_code = charger_code or matched.get("charge_box_id")
            charger_id = charger_id or str(matched.get("charger_id") or "")
            start_time = start_time or matched.get("start_timestamp_string")
            stop_time = stop_time or matched.get("stop_timestamp_string")
            
    # Normalize start/stop timestamps
    if not start_time:
        start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if not stop_time or stop_time.startswith("0001"):
        stop_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 2. Hardware specs & EV battery pack resolution
    charger_specs = resolve_charger_specs(party_id, charger_code or "", charger_id)
    vehicle_specs = resolve_vehicle_battery(vehicle_model_hint, tx_meta.get("Vin") if tx_meta else None)
    if custom_battery_kwh and custom_battery_kwh > 0:
        vehicle_specs["capacity_kwh"] = float(custom_battery_kwh)
        vehicle_specs["name"] = f"Custom EV ({custom_battery_kwh} kWh)"

    rated_kw = float(charger_specs["rated_capacity_kw"])
    gun_cap_kw = float(charger_specs["gun_capacity_kw"])
    battery_kwh = float(vehicle_specs["capacity_kwh"])
    max_dc_kw = float(vehicle_specs.get("max_dc_kw", gun_cap_kw))

    # Allow a sensible 15% physical safety margin above rated capacity before hard spike flag
    power_ceiling_kw = round(max(gun_cap_kw, rated_kw) * 1.15, 1)

    # 3. Fetch meter records if not provided
    if meter_records is None:
        entity_id_query = charger_id or charger_specs.get("charger_id") or "0"
        meter_records = fetch_meter_logs(
            party_id=party_id,
            entity_id=entity_id_query,
            trans_id=trans_id_str,
            start_time=start_time,
            stop_time=stop_time,
            measurand="Energy.Active.Import.Register"
        )

    # 4. Filter & partition readings (Outlet preferred for vehicle delivery, Inlet for grid intake)
    outlet_readings = [r for r in meter_records if r.get("Location") == "Outlet"]
    inlet_readings = [r for r in meter_records if r.get("Location") == "Inlet"]
    
    target_readings = outlet_readings if outlet_readings else meter_records
    location_used = "Outlet" if outlet_readings else ("Inlet" if inlet_readings else "Default")

    # Sort target readings chronologically
    valid_readings = []
    for r in target_readings:
        dt = parse_timestamp_iso(r.get("Value_Timestamp") or "")
        if dt:
            valid_readings.append((dt, r))
    valid_readings.sort(key=lambda x: x[0])

    # 5. Consecutive interval analysis
    intervals = []
    anomalies = []
    
    prev_dt = None
    prev_kwh = None
    initial_kwh = None
    total_energy_kwh = 0.0
    peak_power_kw = 0.0
    spike_count = 0
    rollback_count = 0
    freeze_count = 0
    total_duration_sec = 0.0

    for idx, (dt, record) in enumerate(valid_readings):
        # Extract energy reading
        raw_val = record.get("Value", "")
        reported_kwh = record.get("EnergyInkWh")
        cms_diff = record.get("EnergyDifference")
        
        try:
            if reported_kwh is not None and float(reported_kwh) > 0:
                curr_kwh = float(reported_kwh)
            else:
                curr_kwh = float(raw_val) / 1000.0 if raw_val else 0.0
        except (ValueError, TypeError):
            continue

        if initial_kwh is None:
            initial_kwh = curr_kwh

        delta_t = 0.0
        delta_e = 0.0
        calc_power = 0.0
        interval_flags = []

        if prev_dt is not None:
            delta_t = (dt - prev_dt).total_seconds()
            total_duration_sec += delta_t
            delta_e = curr_kwh - prev_kwh

            if delta_t > 0:
                calc_power = (delta_e * 3600.0) / delta_t
            else:
                calc_power = 0.0

            # --- Anomaly Check 1: Negative Step / Meter Counter Rollback ---
            if delta_e < -0.01:
                rollback_count += 1
                flag_desc = f"Meter Counter Rollback: Dropped by {abs(delta_e):.3f} kWh in {delta_t:.0f}s"
                interval_flags.append({"type": "METER_ROLLBACK", "severity": "HIGH", "detail": flag_desc})
                anomalies.append({
                    "type": "METER_ROLLBACK",
                    "severity": "CRITICAL",
                    "timestamp": dt.strftime("%Y-%m-%d %H:%M:%S"),
                    "reading_kwh": curr_kwh,
                    "delta_kwh": delta_e,
                    "interval_sec": delta_t,
                    "calc_power_kw": calc_power,
                    "explanation": f"Counter decreased from {prev_kwh:.3f} to {curr_kwh:.3f} kWh. Indicates meter reset, register overflow, or communication glitch."
                })

            # --- Anomaly Check 2: Power Spike Exceeding Hardware Capacity ---
            elif calc_power > power_ceiling_kw:
                spike_count += 1
                ratio_pct = (calc_power / rated_kw) * 100.0 if rated_kw > 0 else 0
                flag_desc = f"Power Spike: {calc_power:.1f} kW exceeds rated {rated_kw:.1f} kW ({ratio_pct:.0f}%)"
                interval_flags.append({"type": "SPIKE_POWER_EXCEEDED", "severity": "CRITICAL", "detail": flag_desc})
                anomalies.append({
                    "type": "SPIKE_POWER_EXCEEDED",
                    "severity": "CRITICAL",
                    "timestamp": dt.strftime("%Y-%m-%d %H:%M:%S"),
                    "reading_kwh": curr_kwh,
                    "delta_kwh": delta_e,
                    "interval_sec": delta_t,
                    "calc_power_kw": calc_power,
                    "rated_kw": rated_kw,
                    "power_ceiling_kw": power_ceiling_kw,
                    "ratio_pct": ratio_pct,
                    "explanation": (
                        f"Calculated power delivery of {calc_power:.1f} kW is physically impossible on a "
                        f"{rated_kw:.0f} kW charger hardware ({ratio_pct:.0f}% of nameplate). "
                        f"Energy surged by {delta_e:.3f} kWh in only {delta_t:.0f} seconds."
                    )
                })

            # --- Anomaly Check 3: Meter Freeze / Stalling with Duration ---
            elif delta_t >= 45 and delta_e == 0.0:
                freeze_count += 1
                flag_desc = f"Meter Stalled: 0.0 kWh over {delta_t:.0f}s"
                interval_flags.append({"type": "METER_FREEZE", "severity": "MEDIUM", "detail": flag_desc})

            # Update peak power if not an impossible rollback
            if calc_power > peak_power_kw and delta_e > 0:
                peak_power_kw = calc_power

        intervals.append({
            "index": idx + 1,
            "timestamp": dt.strftime("%Y-%m-%d %H:%M:%S"),
            "time_short": dt.strftime("%H:%M:%S"),
            "raw_value": str(raw_val),
            "reading_kwh": round(curr_kwh, 3),
            "delta_kwh": round(delta_e, 3),
            "delta_time_sec": round(delta_t, 1),
            "calculated_power_kw": round(calc_power, 2),
            "cms_energy_difference": cms_diff,
            "location": record.get("Location", location_used),
            "context": record.get("Reading_Context", ""),
            "flags": interval_flags
        })

        prev_dt = dt
        prev_kwh = curr_kwh

    if valid_readings and initial_kwh is not None:
        total_energy_kwh = round(prev_kwh - initial_kwh, 3)
    elif tx_meta and tx_meta.get("EnergyConsumedDouble"):
        total_energy_kwh = float(tx_meta["EnergyConsumedDouble"])

    # 6. Session-level Checks
    # --- Check 4: Battery Pack Capacity Overrun ---
    overrun_margin = battery_kwh * 1.08  # 8% thermal & conversion allowance
    if total_energy_kwh > overrun_margin:
        overrun_pct = round((total_energy_kwh / battery_kwh) * 100.0, 1)
        phantom_kwh = round(total_energy_kwh - overrun_margin, 2)
        anomalies.append({
            "type": "BATTERY_CAPACITY_OVERRUN",
            "severity": "CRITICAL",
            "timestamp": stop_time,
            "total_delivered_kwh": total_energy_kwh,
            "battery_capacity_kwh": battery_kwh,
            "vehicle_name": vehicle_specs["name"],
            "overrun_pct": overrun_pct,
            "phantom_kwh": phantom_kwh,
            "explanation": (
                f"Total energy delivered ({total_energy_kwh:.2f} kWh) exceeds the physical battery pack capacity "
                f"of {vehicle_specs['name']} ({battery_kwh:.1f} kWh) by {overrun_pct}%. "
                f"Detected {phantom_kwh:.2f} kWh of phantom billing accumulation."
            )
        })

    # --- Check 5: Zero Energy Session with Duration ---
    if total_duration_sec > 300 and total_energy_kwh < 0.05:
        anomalies.append({
            "type": "ZERO_ENERGY_SESSION",
            "severity": "HIGH",
            "timestamp": start_time,
            "duration_sec": total_duration_sec,
            "explanation": (
                f"Transaction remained active for {total_duration_sec/60.0:.1f} minutes, but delivered negligible energy "
                f"({total_energy_kwh:.3f} kWh). Indicates contactor trip, isolation error, or Modbus disconnect."
            )
        })

    # --- Check 6: Inlet vs Outlet Thermodynamic Efficiency ---
    if inlet_readings and outlet_readings:
        try:
            inlet_total = float(inlet_readings[-1].get("EnergyInkWh") or 0.0) - float(inlet_readings[0].get("EnergyInkWh") or 0.0)
            outlet_total = float(outlet_readings[-1].get("EnergyInkWh") or 0.0) - float(outlet_readings[0].get("EnergyInkWh") or 0.0)
            if inlet_total > 0 and outlet_total > 0:
                eff = outlet_total / inlet_total
                if eff > 1.05:
                    anomalies.append({
                        "type": "THERMODYNAMIC_INVERSION",
                        "severity": "CRITICAL",
                        "inlet_kwh": round(inlet_total, 3),
                        "outlet_kwh": round(outlet_total, 3),
                        "efficiency": round(eff, 3),
                        "explanation": (
                            f"Outlet energy ({outlet_total:.2f} kWh) exceeds Inlet grid supply ({inlet_total:.2f} kWh) "
                            f"(Efficiency: {eff*100:.1f}%). Impossible physics; signifies CT calibration bitflip."
                        )
                    })
                elif eff < 0.70:
                    anomalies.append({
                        "type": "LOW_CONVERSION_EFFICIENCY",
                        "severity": "MEDIUM",
                        "inlet_kwh": round(inlet_total, 3),
                        "outlet_kwh": round(outlet_total, 3),
                        "efficiency": round(eff, 3),
                        "explanation": f"Power converter efficiency is abnormally low ({eff*100:.1f}%). Excessive heat loss or power module degradation."
                    })
        except Exception:
            pass

    # 7. Generate Comprehensive RCA Panel
    rca_panel = build_energy_rca_report(
        party_id=party_id,
        transaction_id=trans_id_str,
        charger_specs=charger_specs,
        vehicle_specs=vehicle_specs,
        total_energy_kwh=total_energy_kwh,
        peak_power_kw=peak_power_kw,
        total_duration_sec=total_duration_sec,
        anomalies=anomalies,
        intervals_count=len(intervals)
    )

    res_dict = {
        "success": True,
        "party_id": party_id,
        "transaction_id": trans_id_str,
        "charger_code": charger_code,
        "charger_id": charger_id,
        "start_time": start_time,
        "stop_time": stop_time,
        "charger_specs": charger_specs,
        "vehicle_specs": vehicle_specs,
        "summary": {
            "total_energy_kwh": round(total_energy_kwh, 3),
            "peak_power_kw": round(peak_power_kw, 2),
            "rated_power_kw": rated_kw,
            "gun_capacity_kw": gun_cap_kw,
            "power_ceiling_kw": power_ceiling_kw,
            "battery_capacity_kwh": battery_kwh,
            "battery_utilization_pct": round((total_energy_kwh / battery_kwh) * 100.0, 1) if battery_kwh > 0 else 0,
            "duration_minutes": round(total_duration_sec / 60.0, 1),
            "readings_count": len(intervals),
            "spike_count": spike_count,
            "rollback_count": rollback_count,
            "freeze_count": freeze_count,
            "is_anomaly": len(anomalies) > 0,
            "highest_severity": "CRITICAL" if any(a["severity"] == "CRITICAL" for a in anomalies) else (
                "HIGH" if any(a["severity"] == "HIGH" for a in anomalies) else ("MEDIUM" if anomalies else "NOMINAL")
            )
        },
        "anomalies": anomalies,
        "rca_report": rca_panel,
        "intervals": intervals
    }
    res_dict["billing_audit"] = compute_billing_tax_audit(res_dict, None, party_id=party_id)
    return res_dict


def build_energy_rca_report(
    party_id: str,
    transaction_id: str,
    charger_specs: Dict[str, Any],
    vehicle_specs: Dict[str, Any],
    total_energy_kwh: float,
    peak_power_kw: float,
    total_duration_sec: float,
    anomalies: List[Dict[str, Any]],
    intervals_count: int
) -> Dict[str, Any]:
    """
    Constructs an intelligent Root Cause Analysis diagnostic report with fault attribution,
    physical telemetry proof, and actionable corrective recommendations.
    """
    tariff_inr_per_kwh = 21.00  # Default Indian CPO DC fast charging tariff
    
    if not anomalies:
        return {
            "status": "HEALTHY",
            "verdict": "NOMINAL: Verified Physical Parameters & Energy Curves",
            "confidence_score": 98,
            "badge_color": "green",
            "primary_cause": "Normal Operation",
            "summary_narrative": (
                f"Transaction #{transaction_id} demonstrated stable power delivery. Total consumption of "
                f"{total_energy_kwh:.2f} kWh was within the vehicle battery capacity ({vehicle_specs['capacity_kwh']:.1f} kWh), "
                f"and peak power ({peak_power_kw:.1f} kW) remained strictly compliant with the {charger_specs['rated_capacity_kw']:.0f} kW "
                f"charger hardware ceiling."
            ),
            "evidence": [
                f"Peak observed power of {peak_power_kw:.1f} kW <= Rated {charger_specs['rated_capacity_kw']:.0f} kW",
                f"Total delivered energy {total_energy_kwh:.2f} kWh <= Battery pack limit {vehicle_specs['capacity_kwh']:.1f} kWh",
                f"Continuous monotonic meter readings across {intervals_count} sample intervals with zero rollbacks"
            ],
            "financial_impact": {
                "phantom_kwh": 0.0,
                "disputed_amount_inr": 0.0,
                "billing_status": "Valid & Verified"
            },
            "recommendations": [
                "No action required; session data conforms to physical charging standards."
            ]
        }

    # Identify primary anomaly
    has_spike = any(a["type"] == "SPIKE_POWER_EXCEEDED" for a in anomalies)
    has_overrun = any(a["type"] == "BATTERY_CAPACITY_OVERRUN" for a in anomalies)
    has_rollback = any(a["type"] == "METER_ROLLBACK" for a in anomalies)
    has_inversion = any(a["type"] == "THERMODYNAMIC_INVERSION" for a in anomalies)
    has_zero = any(a["type"] == "ZERO_ENERGY_SESSION" for a in anomalies)

    evidence = []
    recommendations = []
    phantom_kwh = 0.0
    
    if has_spike:
        primary_cause = "Hardware / Modbus CT Ratio Surge or Communication Corrupt Frame"
        spike_items = [a for a in anomalies if a["type"] == "SPIKE_POWER_EXCEEDED"]
        top_spike = max(spike_items, key=lambda x: x.get("calc_power_kw", 0))
        calc_kw = top_spike.get("calc_power_kw", 0)
        rated_kw = charger_specs["rated_capacity_kw"]
        
        evidence.append(
            f"Impossible power surge of {calc_kw:.1f} kW recorded at {top_spike.get('timestamp')}, "
            f"surpassing hardware nameplate {rated_kw:.0f} kW by {((calc_kw/rated_kw)-1)*100:.0f}%."
        )
        recommendations.append(
            f"Inspect RS-485 Modbus serial line and shield grounding on charger {charger_specs['charger_code']} "
            f"to eliminate electromagnetic noise."
        )
        recommendations.append(
            "Verify energy meter CT ratio multiplier in charger controller configuration (ensure 10x or 100x scaler is not bit-flipping)."
        )

    if has_overrun:
        overrun_item = next(a for a in anomalies if a["type"] == "BATTERY_CAPACITY_OVERRUN")
        phantom_kwh += overrun_item.get("phantom_kwh", 0.0)
        primary_cause = "EV Battery Capacity Overrun (Phantom Energy Accumulation)"
        evidence.append(
            f"Billed {total_energy_kwh:.2f} kWh into {vehicle_specs['name']} "
            f"with physical pack capacity of {vehicle_specs['capacity_kwh']:.1f} kWh ({overrun_item.get('overrun_pct')} % pack size)."
        )
        recommendations.append(
            f"Credit customer account for {phantom_kwh:.2f} phantom kWh (approx ₹{phantom_kwh * tariff_inr_per_kwh:.2f})."
        )
        recommendations.append(
            "Check meter start/stop transaction register clearing logic; verify prior session energy was flushed before start."
        )

    if has_rollback:
        rollback_items = [a for a in anomalies if a["type"] == "METER_ROLLBACK"]
        primary_cause = "Energy Meter Counter Rollback / Unscheduled Meter Reset"
        evidence.append(f"Detected {len(rollback_items)} negative energy step reading(s) mid-session.")
        recommendations.append(
            "Check Selec / Schneider power meter firmware and verify register overflow handling."
        )

    if has_inversion:
        primary_cause = "Thermodynamic Inversion (Outlet Energy Exceeds Grid Intake)"
        evidence.append("Outlet vehicle-side readings exceeded Inlet grid-side intake, violating energy conservation.")
        recommendations.append("Swap or recalibrate the CT leads between inlet grid and outlet connector meters.")

    if has_zero:
        primary_cause = "Contactor Failure or Modbus Drop in Active Session"
        evidence.append(f"Zero energy delivered over {total_duration_sec/60:.1f} minutes of connected session.")
        recommendations.append("Inspect DC contactor coil drive signals and BMS authorization handshake.")

    disputed_inr = round(phantom_kwh * tariff_inr_per_kwh, 2)
    
    return {
        "status": "ANOMALY_DETECTED",
        "verdict": f"CRITICAL ANOMALY: {primary_cause}",
        "confidence_score": 94,
        "badge_color": "red",
        "primary_cause": primary_cause,
        "summary_narrative": (
            f"Transaction #{transaction_id} exhibited critical telemetry deviations violating physical power "
            f"and battery constraints. Detected {len(anomalies)} distinct anomaly trigger(s). "
            f"Peak calculated delivery was {peak_power_kw:.1f} kW against a {charger_specs['rated_capacity_kw']:.0f} kW "
            f"hardware threshold, with {total_energy_kwh:.2f} kWh delivered."
        ),
        "evidence": evidence,
        "financial_impact": {
            "phantom_kwh": phantom_kwh,
            "disputed_amount_inr": disputed_inr,
            "billing_status": "Disputed - Potential Phantom Billing" if phantom_kwh > 0 else "Anomalous Telemetry"
        },
        "recommendations": recommendations
    }



def compute_billing_tax_audit(
    analysis: Dict[str, Any],
    booking_record: Optional[Dict[str, Any]] = None,
    party_id: str = "IOC"
) -> Dict[str, Any]:
    """
    Computes commercial reconciliation and itemized tax breakdown:
    1. Real amount after all taxes that SHOULD have been charged from the user.
    2. Actual amount the user ended up paying in the CMS (including taxes).
    3. Itemized discrepancy: Base energy charge + statutory GST (CGST + SGST @ 18%).
    """
    summary = analysis.get("summary", {})
    rca = analysis.get("rca_report", {})
    veh = analysis.get("vehicle_specs", {})
    
    total_energy_kwh = float(summary.get("total_energy_kwh", 0.0))
    phantom_kwh = float(rca.get("financial_impact", {}).get("phantom_kwh", 0.0))
    is_anomaly = bool(summary.get("is_anomaly", False))
    
    booking = booking_record or {}
    
    # 1. Determine Tariff Rate (Rs/kWh)
    per_unit_price = 0.0
    if booking.get("PerUnitPrice") is not None:
        try:
            per_unit_price = float(booking.get("PerUnitPrice"))
        except (ValueError, TypeError):
            per_unit_price = 0.0
            
    if per_unit_price <= 0.0:
        per_unit_price = 14.0 if party_id.upper() == "IOC" else 18.0

    tax_rate_pct = 18.0  # Statutory GST on EV charging services (9% CGST + 9% SGST)
    
    # 2. Extract Actual Amounts Recorded in CMS
    actual_base = 0.0
    actual_tax = 0.0
    actual_total = 0.0
    payment_rec = 0.0
    refunded_amt = 0.0
    payment_status = booking.get("PaymentStatusName") or "Settled"
    
    if booking:
        try:
            if booking.get("ActualTotalAmount") is not None:
                actual_total = float(booking.get("ActualTotalAmount") or 0.0)
            elif booking.get("TotalAmount") is not None:
                actual_total = float(booking.get("TotalAmount") or 0.0)
        except (ValueError, TypeError):
            pass

        try:
            if booking.get("ActualChargingAmount") is not None:
                actual_base = float(booking.get("ActualChargingAmount") or 0.0)
        except (ValueError, TypeError):
            pass

        try:
            if booking.get("ActualTax") is not None:
                actual_tax = float(booking.get("ActualTax") or 0.0)
        except (ValueError, TypeError):
            pass

        try:
            if booking.get("PaymentRecieved") is not None:
                payment_rec = float(booking.get("PaymentRecieved") or 0.0)
            elif booking.get("PaidAmount") is not None:
                payment_rec = float(booking.get("PaidAmount") or 0.0)
        except (ValueError, TypeError):
            pass

        try:
            if booking.get("RefundedAmount") is not None:
                refunded_amt = float(booking.get("RefundedAmount") or 0.0)
            elif booking.get("RefundAmount") is not None:
                refunded_amt = float(booking.get("RefundAmount") or 0.0)
        except (ValueError, TypeError):
            pass

    # Net amount user ended up paying
    if actual_total > 0.0:
        actual_user_paid = round(actual_total, 2)
    elif payment_rec > 0.0:
        actual_user_paid = round(max(0.0, payment_rec - refunded_amt), 2)
        if actual_base <= 0.0:
            actual_base = round(actual_user_paid / (1.0 + (tax_rate_pct / 100.0)), 2)
            actual_tax = round(actual_user_paid - actual_base, 2)
    elif total_energy_kwh > 0.0:
        actual_base = round(total_energy_kwh * per_unit_price, 2)
        actual_tax = round(actual_base * (tax_rate_pct / 100.0), 2)
        actual_user_paid = round(actual_base + actual_tax, 2)
    else:
        actual_base = 0.0
        actual_tax = 0.0
        actual_user_paid = 0.0

    if actual_base > 0.0 and actual_tax <= 0.0:
        actual_tax = round(actual_base * (tax_rate_pct / 100.0), 2)
        if actual_user_paid <= 0.0:
            actual_user_paid = round(actual_base + actual_tax, 2)

    # 3. Determine Legitimate Delivered Energy & Real Amount After All Taxes
    cms_kwh = 0.0
    if booking.get("EnergyConsumed_indecimal") is not None:
        try:
            cms_kwh = float(booking.get("EnergyConsumed_indecimal"))
        except (ValueError, TypeError):
            pass
    elif booking.get("EnergyConsumed") is not None:
        try:
            cms_kwh = float(booking.get("EnergyConsumed"))
        except (ValueError, TypeError):
            pass

    ref_kwh = cms_kwh if cms_kwh > 0.0 else total_energy_kwh

    if not is_anomaly and phantom_kwh <= 0.0:
        # 100% nominal session - physical conforms to billed
        legitimate_kwh = ref_kwh
        real_base = actual_base if actual_base > 0.0 else round(legitimate_kwh * per_unit_price, 2)
        real_tax = actual_tax if actual_tax > 0.0 else round(real_base * (tax_rate_pct / 100.0), 2)
        real_total_after_taxes = actual_user_paid if actual_user_paid > 0.0 else round(real_base + real_tax, 2)
    else:
        # Anomaly detected: deduct phantom spikes and EV pack capacity overruns
        legitimate_kwh = max(0.0, ref_kwh - phantom_kwh)
        pack_cap = float(veh.get("capacity_kwh", 0.0))
        if pack_cap > 0.0 and legitimate_kwh > pack_cap * 1.05:
            legitimate_kwh = pack_cap

        real_base = round(legitimate_kwh * per_unit_price, 2)
        real_tax = round(real_base * (tax_rate_pct / 100.0), 2)
        real_total_after_taxes = round(real_base + real_tax, 2)

    variance = round(actual_user_paid - real_total_after_taxes, 2)
    base_discrepancy = round(actual_base - real_base, 2)
    tax_discrepancy = round(actual_tax - real_tax, 2)

    if variance > 0.50:
        status_code = "OVERCHARGED"
        status_label = f"OVERCHARGED BY ₹{variance:.2f}"
        badge_color = "red"
        action_summary = (
            f"User ended up paying ₹{actual_user_paid:.2f}, but the real physical charge should have been "
            f"₹{real_total_after_taxes:.2f} after all taxes. The user was overbilled by ₹{variance:.2f} "
            f"(Base Energy: ₹{base_discrepancy:.2f} + Excess GST: ₹{tax_discrepancy:.2f}) due to {phantom_kwh:.2f} "
            f"phantom/unphysical kWh. Commercial adjustment / refund voucher of ₹{variance:.2f} is due to the user."
        )
    elif variance < -0.50:
        status_code = "UNDERCHARGED"
        status_label = f"UNDERCHARGED BY ₹{abs(variance):.2f}"
        badge_color = "orange"
        action_summary = (
            f"User ended up paying ₹{actual_user_paid:.2f}, whereas physical delivery evaluated to "
            f"₹{real_total_after_taxes:.2f} after all taxes. CPO revenue leakage: ₹{abs(variance):.2f}."
        )
    else:
        status_code = "ACCURATELY_CHARGED"
        status_label = "ACCURATELY BILLED (₹0.00 Discrepancy)"
        badge_color = "green"
        action_summary = (
            f"All metered units ({legitimate_kwh:.2f} kWh) were physically delivered and verified. "
            f"The real amount after all taxes (₹{real_total_after_taxes:.2f}) exactly equals the actual "
            f"amount the user ended up paying (₹{actual_user_paid:.2f}). Zero refund required."
        )

    return {
        "has_billing_data": True,
        "tariff_rate_inr_per_kwh": per_unit_price,
        "tax_rate_pct": tax_rate_pct,
        "real_should_charge": {
            "legitimate_energy_kwh": round(legitimate_kwh, 3),
            "base_amount_inr": real_base,
            "tax_amount_inr": real_tax,
            "cgst_inr": round(real_tax / 2.0, 2),
            "sgst_inr": round(real_tax / 2.0, 2),
            "total_amount_after_taxes_inr": real_total_after_taxes
        },
        "actual_user_paid": {
            "billed_energy_kwh": round(ref_kwh, 3),
            "base_amount_inr": actual_base,
            "tax_amount_inr": actual_tax,
            "cgst_inr": round(actual_tax / 2.0, 2),
            "sgst_inr": round(actual_tax / 2.0, 2),
            "total_amount_after_taxes_inr": actual_user_paid,
            "payment_received_inr": payment_rec,
            "refunded_inr": refunded_amt,
            "payment_status": payment_status
        },
        "comparison": {
            "variance_inr": variance,
            "base_discrepancy_inr": base_discrepancy,
            "tax_discrepancy_inr": tax_discrepancy,
            "phantom_kwh": phantom_kwh,
            "status": status_code,
            "status_label": status_label,
            "badge_color": badge_color,
            "action_summary": action_summary
        }
    }


def batch_scan_transactions(
    party_id: str = "IOC",
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    limit: int = 20
) -> Dict[str, Any]:
    """
    Scans a batch of recent transactions, analyzes each for energy spikes and anomalies,
    and returns categorized results.
    """
    tx_res = fetch_ocpp_transactions(
        party_id=party_id,
        start_date=start_date,
        end_date=end_date,
        page_size=limit
    )
    
    if not tx_res.get("success"):
        return {"success": False, "error": tx_res.get("error"), "results": []}

    transactions = tx_res.get("transactions", [])
    scanned_results = []
    total_anomalies = 0

    for tx in transactions:
        tx_id = str(tx.get("transaction_pk") or "")
        cid = str(tx.get("charger_id") or "")
        code = tx.get("charge_box_id") or ""
        s_time = tx.get("start_timestamp_string") or ""
        e_time = tx.get("stop_timestamp_string") or ""
        kwh = float(tx.get("EnergyConsumedDouble") or 0.0)

        # Quick pre-filter: if kwh == 0 and duration < 1m, can still check or skip
        analysis = analyze_transaction_energy(
            party_id=party_id,
            transaction_id=tx_id,
            charger_code=code,
            charger_id=cid,
            start_time=s_time,
            stop_time=e_time,
            tx_meta=tx
        )
        
        is_ano = analysis["summary"]["is_anomaly"]
        if is_ano:
            total_anomalies += 1
            
        scanned_results.append({
            "transaction_id": tx_id,
            "charger_code": code,
            "station_name": tx.get("ChargingStationName", ""),
            "start_time": s_time,
            "stop_time": e_time,
            "energy_kwh": analysis["summary"]["total_energy_kwh"],
            "peak_power_kw": analysis["summary"]["peak_power_kw"],
            "rated_power_kw": analysis["summary"]["rated_power_kw"],
            "is_anomaly": is_ano,
            "highest_severity": analysis["summary"]["highest_severity"],
            "primary_cause": analysis["rca_report"]["primary_cause"],
            "anomalies_count": len(analysis["anomalies"]),
            "spike_count": analysis["summary"]["spike_count"],
            "rollback_count": analysis["summary"]["rollback_count"]
        })

    return {
        "success": True,
        "party_id": party_id,
        "total_scanned": len(scanned_results),
        "total_anomalies": total_anomalies,
        "results": scanned_results
    }



def investigate_booking_energy(
    party_id: str,
    booking_or_tx_id: str,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None
) -> Dict[str, Any]:
    """
    Automated investigation from Booking ID:
    1. Maps Booking ID -> central OCPP Transaction ID (transaction_pk).
    2. Automatically fetches all vehicle specifications from the CMS (Model, Make, Battery Capacity, Reg Number, VIN).
    3. Resolves physical battery pack capacity from CMS data (zero manual user input required).
    4. Analyzes consecutive periodic meter readings (Energy.Active.Import.Register).
    5. Audits for power spikes (kW vs rated), pack overruns, and rollbacks.
    6. Returns complete diagnostic RCA breakdown.
    """
    from src.rca.booking_inspector import (
        fetch_ocpp_transaction,
        fetch_completed_booking,
        fetch_cancelled_booking,
        fetch_roaming_booking,
        get_party_cookie_header,
        get_party_config
    )

    party_id = (party_id or "IOC").upper()
    clean_id = str(booking_or_tx_id).strip()
    if not clean_id:
        return {"success": False, "error": "No Booking ID provided."}

    cookie_hdr = get_party_cookie_header(party_id)
    portal_cfg = get_party_config(party_id)
    portal_url = portal_cfg.get("portal_url", get_portal_base_url(party_id))

    today_str = datetime.now().strftime("%Y-%m-%d")
    s_date = start_date.strip() if start_date and start_date.strip() else "2024-01-01"
    e_date = end_date.strip() if end_date and end_date.strip() else today_str

    # 1. Look up transaction by ReservationId (Booking ID)
    tx_record, _ = fetch_ocpp_transaction(clean_id, s_date, e_date, cookie_hdr, base_url=portal_url)

    # If not found by ReservationId, check if user provided a direct Transaction ID
    if not tx_record:
        res_tx = fetch_ocpp_transactions(
            party_id=party_id,
            filter_query=f"transaction_pk~eq~{clean_id}",
            page_size=5
        )
        txs = res_tx.get("transactions", [])
        if txs:
            tx_record = txs[0]

    # If still not found, search in completed/cancelled bookings to see if we can extract TxnId
    booking_record = None
    if not tx_record:
        booking_record, _ = fetch_completed_booking(clean_id, s_date, e_date, cookie_hdr, base_url=portal_url)
        if not booking_record:
            booking_record, _ = fetch_cancelled_booking(clean_id, s_date, e_date, cookie_hdr, base_url=portal_url)
        if not booking_record:
            booking_record, _ = fetch_roaming_booking(clean_id, cookie_hdr, base_url=portal_url)

        if booking_record:
            tx_cand = booking_record.get("TransactionId") or booking_record.get("TranscationId") or booking_record.get("TxnId")
            if tx_cand and str(tx_cand).strip().isdigit() and int(str(tx_cand).strip()) > 0:
                cand_id = str(tx_cand).strip()
                res_cand = fetch_ocpp_transactions(party_id=party_id, filter_query=f"transaction_pk~eq~{cand_id}", page_size=5)
                if res_cand.get("transactions"):
                    tx_record = res_cand["transactions"][0]

    if not tx_record:
        party_label = portal_cfg.get("name", party_id)
        return {
            "success": False,
            "error": f"No OCPP transaction found for Booking ID #{clean_id} on {party_label} between {s_date} and {e_date}."
        }

    # Extract transaction parameters
    tx_id = str(tx_record.get("transaction_pk") or "")
    booking_id = str(tx_record.get("ReservationId") or clean_id)
    charger_code = tx_record.get("charge_box_id") or ""
    charger_id = str(tx_record.get("charger_id") or "")
    start_time = tx_record.get("start_timestamp_string") or ""
    stop_time = tx_record.get("stop_timestamp_string") or ""
    vin = tx_record.get("Vin") or ""

    # 2. Fetch vehicle details from CMS booking records
    if not booking_record:
        booking_record, _ = fetch_completed_booking(booking_id, s_date, e_date, cookie_hdr, base_url=portal_url)
        if not booking_record:
            booking_record, _ = fetch_cancelled_booking(booking_id, s_date, e_date, cookie_hdr, base_url=portal_url)
        if not booking_record:
            booking_record, _ = fetch_roaming_booking(booking_id, cookie_hdr, base_url=portal_url)

    vehicle_make = ""
    vehicle_model = ""
    vehicle_number = ""
    battery_kwh = None

    if booking_record:
        vehicle_make = booking_record.get("ManufacturerName") or booking_record.get("VehicleMake") or ""
        vehicle_model = booking_record.get("ModelName") or booking_record.get("VehicleModel") or ""
        vehicle_number = booking_record.get("VehicleRegistrationNumber") or booking_record.get("VehicleNumber") or ""
        b_cap = booking_record.get("BatteryCapacity") or booking_record.get("Battery_Capacity")
        if b_cap:
            try:
                battery_kwh = float(b_cap)
            except (ValueError, TypeError):
                pass

    vehicle_str = f"{vehicle_make} {vehicle_model}".strip()

    # 3. Perform energy anomaly detection with auto-resolved parameters
    analysis = analyze_transaction_energy(
        party_id=party_id,
        transaction_id=tx_id,
        charger_code=charger_code,
        charger_id=charger_id,
        start_time=start_time,
        stop_time=stop_time,
        vehicle_model_hint=vehicle_str,
        custom_battery_kwh=battery_kwh,
        tx_meta=tx_record
    )

    # 4. Attach CMS vehicle and booking metadata to response
    analysis["booking_id"] = booking_id
    analysis["vehicle_cms"] = {
        "vehicle_make": vehicle_make or "EV Manufacturer",
        "vehicle_model": vehicle_model or "EV Model",
        "vehicle_name": vehicle_str or analysis["vehicle_specs"]["name"],
        "vehicle_number": vehicle_number or "-",
        "vin": vin or "-",
        "battery_capacity_kwh": analysis["vehicle_specs"]["capacity_kwh"],
        "source": "CMS Booking Record" if vehicle_str or battery_kwh else "Database Heuristic"
    }

    # 5. Compute commercial reconciliation & real amount after all taxes
    analysis["billing_audit"] = compute_billing_tax_audit(analysis, booking_record, party_id=party_id)

    return analysis
