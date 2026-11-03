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
from datetime import datetime, timedelta
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



def fetch_and_correlate_ocpp_telemetry(
    party_id: str,
    charger_id: str,
    charger_code: str,
    tx_id: str,
    start_time: str,
    stop_time: str,
    booking_record: Optional[Dict[str, Any]] = None,
    tx_record: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Fetches raw OCPP event logs for the charger around session start and stop windows.
    Parses full OCPP message payloads:
      - StartTransaction (meterStart, connectorId, idTag)
      - StatusNotification transitions (e.g. Charging -> Available abort)
      - MeterValues (sampledValue arrays, measurands, units, timestamps)
      - StopTransaction (meterStop, stop_reason)
    Cross-checks against transaction records and CMS billing to detect discrepancies.
    """
    from src.rca.booking_inspector import get_party_cookie_header
    
    party_id = (party_id or "IOC").upper()
    portal_base = get_portal_base_url(party_id)
    cookie_hdr = get_party_cookie_header(party_id)
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "User-Agent": DEFAULT_BROWSER_UA,
        "Cookie": cookie_hdr
    }
    
    def parse_dt(s: str) -> Optional[datetime]:
        if not s or s.startswith("0001"):
            return None
        clean = s.split(".")[0].replace("T", " ")
        try:
            return datetime.strptime(clean, "%Y-%m-%d %H:%M:%S")
        except Exception:
            return None

    dt_start = parse_dt(start_time)
    dt_stop = parse_dt(stop_time)

    def fetch_logs(f_date: str, t_date: str) -> List[Dict[str, Any]]:
        url = (
            f"{portal_base}/GetOcppLogs"
            f"?entityId={charger_id}"
            f"&fromDate={urllib.parse.quote(f_date)}"
            f"&toDate={urllib.parse.quote(t_date)}"
        )
        payload = urllib.parse.urlencode({
            "page": "1",
            "pageSize": "100",
            "sort": "Date-desc",
            "filter": ""
        }).encode("utf-8")
        req = urllib.request.Request(url, data=payload, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=12) as res:
                return json.loads(res.read().decode("utf-8")).get("Data", [])
        except Exception:
            return []

    start_events = []
    if dt_start:
        w_start_from = (dt_start - timedelta(minutes=5)).strftime("%Y-%m-%d %H:%M:%S")
        w_start_to = (dt_start + timedelta(minutes=15)).strftime("%Y-%m-%d %H:%M:%S")
        start_events = fetch_logs(w_start_from, w_start_to)

    stop_events = []
    if dt_stop and dt_start and (dt_stop - dt_start).total_seconds() > 900:
        w_stop_from = (dt_stop - timedelta(minutes=15)).strftime("%Y-%m-%d %H:%M:%S")
        w_stop_to = (dt_stop + timedelta(minutes=5)).strftime("%Y-%m-%d %H:%M:%S")
        stop_events = fetch_logs(w_stop_from, w_stop_to)

    all_raw_events = list(reversed(start_events)) + list(reversed(stop_events))
    
    start_tx_info = None
    stop_tx_info = None
    meter_values_list = []
    status_transitions = []
    other_tx_ids = set()

    for it in all_raw_events:
        d = it.get("Date")
        msg = str(it.get("Message") or it.get("LogMessage") or "")
        try:
            arr = json.loads(msg)
            if isinstance(arr, list) and len(arr) >= 3 and arr[0] == 2:
                action = arr[2]
                payload = arr[3] if len(arr) > 3 else {}
                
                if action == "StartTransaction":
                    s_meter = payload.get("meterStart")
                    tag = payload.get("idTag")
                    conn = payload.get("connectorId")
                    if str(payload.get("reservationId")) == str(tx_id) or not start_tx_info:
                        start_tx_info = {
                            "timestamp": d,
                            "meter_start_wh": s_meter,
                            "meter_start_kwh": (float(s_meter) / 1000.0) if s_meter is not None else None,
                            "id_tag": tag,
                            "connector_id": conn,
                            "raw": payload
                        }
                elif action == "StopTransaction":
                    t_id = str(payload.get("transactionId"))
                    if t_id == str(tx_id):
                        stop_tx_info = {
                            "timestamp": d,
                            "transaction_id": t_id,
                            "meter_stop_wh": payload.get("meterStop"),
                            "meter_stop_kwh": (float(payload.get("meterStop")) / 1000.0) if payload.get("meterStop") is not None else None,
                            "reason": payload.get("reason"),
                            "raw": payload
                        }
                    else:
                        other_tx_ids.add(t_id)
                elif action == "MeterValues":
                    t_id = str(payload.get("transactionId"))
                    if t_id == str(tx_id):
                        m_vals = payload.get("meterValue", [])
                        samples_parsed = []
                        for mv in m_vals:
                            mv_ts = mv.get("timestamp")
                            for s in mv.get("sampledValue", []):
                                samples_parsed.append({
                                    "timestamp": mv_ts or d,
                                    "value": s.get("value"),
                                    "measurand": s.get("measurand", "Energy.Active.Import.Register"),
                                    "unit": s.get("unit", "Wh"),
                                    "context": s.get("context", "")
                                })
                        meter_values_list.append({
                            "timestamp": d,
                            "transaction_id": t_id,
                            "samples": samples_parsed,
                            "raw": payload
                        })
                    else:
                        other_tx_ids.add(t_id)
                elif action == "StatusNotification":
                    status_transitions.append({
                        "timestamp": d,
                        "connector_id": payload.get("connectorId"),
                        "status": payload.get("status"),
                        "error_code": payload.get("errorCode")
                    })
        except Exception:
            pass

    # Discrepancy analysis
    discrepancies = []
    
    # 1. Immediate connector status abort
    if status_transitions:
        chg_idx = next((i for i, s in enumerate(status_transitions) if s["status"] == "Charging"), None)
        if chg_idx is not None and chg_idx + 1 < len(status_transitions):
            next_st = status_transitions[chg_idx + 1]
            if next_st["status"] in ("Available", "Faulted", "SuspendedEV"):
                discrepancies.append({
                    "severity": "CRITICAL",
                    "category": "CONNECTOR_STATE_ABORT",
                    "title": "Connector Abruptly Released to Available at Session Start",
                    "detail": f"Connector entered 'Charging' at {status_transitions[chg_idx]['timestamp']}, but immediately reverted to '{next_st['status']}' at {next_st['timestamp']}. No actual energy transfer occurred."
                })

    # 2. MeterValues stream missing / silent
    if len(meter_values_list) <= 1:
        discrepancies.append({
            "severity": "HIGH",
            "category": "METER_VALUES_STREAM_DROPPED",
            "title": f"Telemetry Stream Dropped ({len(meter_values_list)} MeterValues Payload Captured)",
            "detail": f"Only {len(meter_values_list)} MeterValues payload was received from the charger. Periodic telemetry was not transmitted for the session duration."
        })

    # 3. Orphaned / Zombie transaction with subsequent interleaved sessions
    if dt_start and dt_stop:
        dur_hours = (dt_stop - dt_start).total_seconds() / 3600.0
        if dur_hours > 4:
            other_tx_str = ", ".join(list(other_tx_ids)[:5]) if other_tx_ids else "None"
            discrepancies.append({
                "severity": "CRITICAL",
                "category": "ZOMBIE_ORPHANED_TRANSACTION",
                "title": f"Orphaned Zombie Session (Duration: {dur_hours:.1f} Hours / {dur_hours/24:.1f} Days)",
                "detail": (
                    f"Transaction was initiated at {start_time} but StopTransaction was not sent until {stop_time} ({dur_hours:.1f} hours later). "
                    f"Other independent transactions ({other_tx_str}) were initiated and closed on this charger while this transaction remained stuck in CMS."
                )
            })

    # 4. Start vs Stop vs CMS Billed delta discrepancy
    s_val = start_tx_info.get("meter_start_wh") if start_tx_info else (tx_record.get("start_value") if tx_record else None)
    e_val = stop_tx_info.get("meter_stop_wh") if stop_tx_info else (tx_record.get("stop_value") if tx_record else None)
    
    first_mv = None
    last_mv = None
    energy_readings = []
    if meter_values_list:
        for mv in meter_values_list:
            for s in mv.get("samples", []):
                meas = s.get("measurand", "")
                unit = s.get("unit", "")
                # Only extract active energy register values, ignoring Voltage, Current, Power, and SoC
                if meas in ("Energy.Active.Import.Register", "Energy.Active.Import", "") or unit in ("Wh", "kWh"):
                    try:
                        val = float(s.get("value"))
                        wh_val = val if (unit != "kWh" and val > 1000) else val * 1000.0
                        energy_readings.append(wh_val)
                        break
                    except (ValueError, TypeError):
                        continue

    if energy_readings:
        first_mv = energy_readings[0]
        last_mv = energy_readings[-1]
        actual_mv_delivered = max(0.0, (last_mv - first_mv) / 1000.0)
    else:
        actual_mv_delivered = 0.0

    # Cross-reference against verified total metered energy in CMS records
    reported_consumed_kwh = None
    if tx_record and tx_record.get("EnergyConsumedDouble") is not None:
        try:
            reported_consumed_kwh = float(tx_record["EnergyConsumedDouble"])
        except (ValueError, TypeError):
            pass
    elif booking_record and booking_record.get("EnergyConsumed_indecimal") is not None:
        try:
            reported_consumed_kwh = float(booking_record["EnergyConsumed_indecimal"])
        except (ValueError, TypeError):
            pass

    if reported_consumed_kwh is not None and reported_consumed_kwh > actual_mv_delivered:
        actual_mv_delivered = reported_consumed_kwh

    if s_val is not None and e_val is not None:
        try:
            raw_delta_kwh = (float(e_val) - float(s_val)) / 1000.0
            gap_kwh = raw_delta_kwh - actual_mv_delivered
            
            # Material discrepancy check: must exceed 0.50 kWh and 2% variance to avoid flagging minor buffer flushes
            if gap_kwh > 0.50 and (actual_mv_delivered < 0.05 or (gap_kwh / max(0.1, raw_delta_kwh)) > 0.02):
                if actual_mv_delivered < 0.05:
                    detail_text = (
                        f"StopTransaction recorded meterStop: {e_val} Wh vs meterStart: {s_val} Wh (Delta: {raw_delta_kwh:.2f} kWh). "
                        f"However, zero physical energy was delivered to the vehicle during this session. "
                        f"The entire billed amount ({raw_delta_kwh:.2f} kWh) is phantom energy from an aborted session."
                    )
                else:
                    detail_text = (
                        f"StopTransaction recorded meterStop: {e_val} Wh vs meterStart: {s_val} Wh (Delta: {raw_delta_kwh:.2f} kWh). "
                        f"However, live metered delivery during the active charging session accounted for {actual_mv_delivered:.3f} kWh. "
                        f"The unmetered difference (+{gap_kwh:.2f} kWh) was billed in excess of physical delivery."
                    )
                discrepancies.append({
                    "severity": "CRITICAL",
                    "category": "STOP_METER_LEAK_DISCREPANCY",
                    "title": f"CMS Billed {raw_delta_kwh:.2f} kWh vs {actual_mv_delivered:.3f} kWh Actually Metered",
                    "detail": detail_text,
                    "phantom_kwh": round(gap_kwh, 3)
                })
        except Exception:
            pass

    return {
        "charger_code": charger_code,
        "charger_id": charger_id,
        "transaction_id": tx_id,
        "start_event": start_tx_info,
        "stop_event": stop_tx_info,
        "meter_values_count": len(meter_values_list),
        "meter_values_payloads": meter_values_list,
        "status_transitions": status_transitions,
        "other_transactions_detected": list(other_tx_ids),
        "discrepancies": discrepancies,
        "is_zombie_session": any(d["category"] == "ZOMBIE_ORPHANED_TRANSACTION" for d in discrepancies),
        "is_zero_energy_abort": any(d["category"] == "CONNECTOR_STATE_ABORT" for d in discrepancies) or (len(meter_values_list) <= 1 and actual_mv_delivered < 0.05)
    }


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
    tx_meta: Optional[Dict[str, Any]] = None,
    tariff_per_unit: Optional[float] = None,
    booking_record: Optional[Dict[str, Any]] = None,
    ocpp_telemetry: Optional[Dict[str, Any]] = None
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

    # Connector unitary capacity: Gun capacity takes precedence as each connector is physically bounded
    connector_kw = gun_cap_kw if gun_cap_kw > 0 else (rated_kw if rated_kw > 0 else 30.0)
    # Unitary transfer rates:
    # 1 hour = connector_kw kWh
    # 1 minute = connector_kw / 60.0 kWh
    # 1 second = connector_kw / 3600.0 kWh
    rate_per_sec = connector_kw / 3600.0
    rate_per_min = connector_kw / 60.0

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
        expected_energy_kwh = 0.0
        max_allowed_energy_kwh = 0.0
        interval_flags = []

        if prev_dt is not None:
            delta_t = (dt - prev_dt).total_seconds()
            total_duration_sec += delta_t
            delta_e = curr_kwh - prev_kwh

            # Unitary energy transfer comparison grounded in connector physical limit
            # A connector rated at P kW transfers:
            # - P kWh in 1 hour (3600 seconds)
            # - P / 60 kWh in 1 minute (60 seconds)
            # - P / 3600 kWh in 1 second
            if delta_t > 0:
                expected_energy_kwh = rate_per_sec * delta_t
                max_allowed_energy_kwh = expected_energy_kwh * 1.15  # 15% physical tolerance for sampling jitter
                calc_power = (delta_e * 3600.0) / delta_t  # Informational only, never used for anomaly gating
            else:
                expected_energy_kwh = 0.0
                max_allowed_energy_kwh = 0.005
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

            # --- Anomaly Check 2: Unitary Energy Transfer Limit Exceeded (Energy Difference vs Duration) ---
            elif delta_e > max_allowed_energy_kwh and delta_e > 0.005:
                spike_count += 1
                ratio_pct = (delta_e / expected_energy_kwh) * 100.0 if expected_energy_kwh > 0 else 9999.0
                flag_desc = (
                    f"Abnormal Energy Jump: +{delta_e:.3f} kWh in {delta_t:.0f}s exceeds "
                    f"{connector_kw:.0f} kW connector limit (max {max_allowed_energy_kwh:.3f} kWh)"
                )
                interval_flags.append({"type": "ABNORMAL_ENERGY_JUMP", "severity": "CRITICAL", "detail": flag_desc})
                anomalies.append({
                    "type": "ABNORMAL_ENERGY_JUMP",
                    "severity": "CRITICAL",
                    "timestamp": dt.strftime("%Y-%m-%d %H:%M:%S"),
                    "reading_kwh": curr_kwh,
                    "delta_kwh": delta_e,
                    "interval_sec": delta_t,
                    "connector_kw": connector_kw,
                    "rate_per_sec": round(rate_per_sec, 5),
                    "rate_per_min": round(rate_per_min, 3),
                    "expected_kwh": round(expected_energy_kwh, 4),
                    "max_allowed_kwh": round(max_allowed_energy_kwh, 4),
                    "ratio_pct": round(ratio_pct, 1),
                    "calc_power_kw": round(calc_power, 2),
                    "explanation": (
                        f"Energy increment of +{delta_e:.3f} kWh in {delta_t:.0f}s is physically impossible for a "
                        f"{connector_kw:.0f} kW connector. A {connector_kw:.0f} kW connector delivers {connector_kw:.0f} kWh in an hour "
                        f"({rate_per_min:.2f} kWh/min, {rate_per_sec:.5f} kWh/s). Maximum physical transfer in {delta_t:.0f}s is {expected_energy_kwh:.3f} kWh. "
                        f"Reported energy difference exceeds physical capacity by {ratio_pct:.0f}%."
                    )
                })

            # --- Anomaly Check 3: Meter Freeze / Stalling with Duration ---
            elif delta_t >= 45 and delta_e == 0.0:
                freeze_count += 1
                flag_desc = f"Meter Stalled: 0.0 kWh over {delta_t:.0f}s"
                interval_flags.append({"type": "METER_FREEZE", "severity": "MEDIUM", "detail": flag_desc})

            # Informational peak power tracking only
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
            "expected_max_kwh": round(expected_energy_kwh, 4) if prev_dt is not None else 0.0,
            "max_allowed_kwh": round(max_allowed_energy_kwh, 4) if prev_dt is not None else 0.0,
            "rate_per_sec_kwh": round(rate_per_sec, 5) if prev_dt is not None else 0.0,
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

    # Extract CMS recorded energy for cross-check
    cms_tx_kwh = 0.0
    if booking_record:
        if booking_record.get("EnergyConsumed_indecimal") is not None:
            try:
                cms_tx_kwh = float(booking_record["EnergyConsumed_indecimal"])
            except (ValueError, TypeError):
                pass
        elif booking_record.get("EnergyConsumed") is not None:
            try:
                cms_tx_kwh = float(booking_record["EnergyConsumed"])
            except (ValueError, TypeError):
                pass
    if cms_tx_kwh <= 0 and tx_meta:
        if tx_meta.get("EnergyConsumedDouble") is not None:
            try:
                cms_tx_kwh = float(tx_meta["EnergyConsumedDouble"])
            except (ValueError, TypeError):
                pass
        elif tx_meta.get("TotalMeterReading") is not None:
            try:
                cms_tx_kwh = float(tx_meta["TotalMeterReading"]) / 1000.0
            except (ValueError, TypeError):
                pass
        elif tx_meta.get("stop_value") and tx_meta.get("start_value"):
            try:
                s_val = float(tx_meta["start_value"])
                e_val = float(tx_meta["stop_value"])
                if e_val > s_val:
                    cms_tx_kwh = (e_val - s_val) / 1000.0
            except (ValueError, TypeError):
                pass

    dt_s = parse_timestamp_iso(start_time)
    dt_e = parse_timestamp_iso(stop_time)
    session_duration_sec = total_duration_sec
    if dt_s and dt_e and dt_e > dt_s:
        session_duration_sec = (dt_e - dt_s).total_seconds()

    # 6. Session-level Checks
    # --- Check 4: Discrepancy between CMS Billed Energy vs Physical Meter Readings ---
    # Only flag if discrepancy is material (> 0.5 kWh and > 2% of billed, or total_energy_kwh == 0)
    gap_diff = cms_tx_kwh - total_energy_kwh
    if (gap_diff > 0.50 and (gap_diff / max(0.1, cms_tx_kwh)) > 0.02) or (total_energy_kwh < 0.05 and cms_tx_kwh > 0.5):
        phantom_diff = round(gap_diff, 3)
        anomalies.append({
            "type": "DISCREPANCY_BILLED_VS_METERED",
            "severity": "CRITICAL",
            "timestamp": stop_time,
            "metered_kwh": total_energy_kwh,
            "billed_kwh": cms_tx_kwh,
            "phantom_kwh": phantom_diff,
            "explanation": (
                f"Physical meter logs show {total_energy_kwh:.3f} kWh delivered to the vehicle, "
                f"yet CMS recorded and billed {cms_tx_kwh:.2f} kWh (Difference: +{phantom_diff:.2f} kWh unmetered phantom energy)."
            )
        })

    # --- Check 5: Battery Pack Capacity Overrun (on either physical metered or CMS billed energy) ---
    effective_overrun_energy = max(total_energy_kwh, cms_tx_kwh)
    overrun_margin = battery_kwh * 1.08  # 8% thermal & conversion allowance
    if effective_overrun_energy > overrun_margin:
        overrun_pct = round((effective_overrun_energy / battery_kwh) * 100.0, 1)
        phantom_kwh = round(effective_overrun_energy - battery_kwh, 2)
        anomalies.append({
            "type": "BATTERY_CAPACITY_OVERRUN",
            "severity": "CRITICAL",
            "timestamp": stop_time,
            "total_delivered_kwh": effective_overrun_energy,
            "battery_capacity_kwh": battery_kwh,
            "vehicle_name": vehicle_specs["name"],
            "overrun_pct": overrun_pct,
            "phantom_kwh": phantom_kwh,
            "explanation": (
                f"Energy recorded ({effective_overrun_energy:.2f} kWh) exceeds physical battery pack capacity "
                f"of {vehicle_specs['name']} ({battery_kwh:.1f} kWh) by {overrun_pct}%. "
                f"Detected {phantom_kwh:.2f} kWh of physically impossible phantom billing accumulation."
            )
        })

    # --- Check 6: Zero Energy Delivered / Session Abort ---
    if total_energy_kwh < 0.05:
        anomalies.append({
            "type": "ZERO_ENERGY_DELIVERY",
            "severity": "CRITICAL" if cms_tx_kwh > 0.1 else "HIGH",
            "timestamp": start_time,
            "duration_sec": session_duration_sec,
            "metered_kwh": total_energy_kwh,
            "billed_kwh": cms_tx_kwh,
            "explanation": (
                f"Zero energy was delivered to the vehicle ({total_energy_kwh:.3f} kWh metered). "
                f"Connector disconnected immediately, contactor failed to close, or session aborted during initialization."
            )
        })

    # --- Check 7: Orphaned / Zombie Session ---
    if session_duration_sec > 14400 and (len(valid_readings) <= 3 or (ocpp_telemetry and ocpp_telemetry.get("is_zombie_session"))):
        dur_h = round(session_duration_sec / 3600.0, 1)
        anomalies.append({
            "type": "ORPHANED_ZOMBIE_TRANSACTION",
            "severity": "CRITICAL",
            "timestamp": stop_time,
            "duration_hours": dur_h,
            "readings_count": len(valid_readings),
            "explanation": (
                f"Transaction remained hung open in CMS for {dur_h:.1f} hours ({dur_h/24:.1f} days) "
                f"with only {len(valid_readings)} meter reading(s). Physical meter was advanced by other charging sessions."
            )
        })

    # Merge OCPP Telemetry discrepancies if present
    if ocpp_telemetry and ocpp_telemetry.get("discrepancies"):
        for d in ocpp_telemetry["discrepancies"]:
            if not any(a.get("type") == d["category"] for a in anomalies):
                anomalies.append({
                    "type": d["category"],
                    "severity": d["severity"],
                    "title": d["title"],
                    "explanation": d["detail"],
                    "timestamp": start_time
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

    # 7. Extract Per Unit Price from Charging Details / tx_meta if not provided
    active_tariff = tariff_per_unit
    if (active_tariff is None or active_tariff <= 0) and tx_meta:
        if tx_meta.get("PerUnitPrice") is not None:
            try:
                active_tariff = float(tx_meta["PerUnitPrice"])
            except (ValueError, TypeError):
                pass

    # Generate Comprehensive RCA Panel with exact Charging Details Per Unit Price
    rca_panel = build_energy_rca_report(
        party_id=party_id,
        transaction_id=trans_id_str,
        charger_specs=charger_specs,
        vehicle_specs=vehicle_specs,
        total_energy_kwh=total_energy_kwh,
        peak_power_kw=peak_power_kw,
        total_duration_sec=session_duration_sec,
        anomalies=anomalies,
        intervals_count=len(intervals),
        tariff_inr_per_kwh=active_tariff,
        ocpp_telemetry=ocpp_telemetry,
        cms_billed_kwh=cms_tx_kwh
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
            "connector_capacity_kw": connector_kw,
            "gun_capacity_kw": gun_cap_kw,
            "rated_power_kw": rated_kw,
            "rate_per_hour_kwh": connector_kw,
            "rate_per_min_kwh": round(rate_per_min, 3),
            "rate_per_sec_kwh": round(rate_per_sec, 5),
            "peak_power_kw": round(peak_power_kw, 2),
            "battery_capacity_kwh": battery_kwh,
            "battery_utilization_pct": round((total_energy_kwh / battery_kwh) * 100.0, 1) if battery_kwh > 0 else 0,
            "duration_minutes": round(total_duration_sec / 60.0, 1),
            "readings_count": len(intervals),
            "spike_count": spike_count,
            "abnormal_jump_count": spike_count,
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
    intervals_count: int,
    tariff_inr_per_kwh: Optional[float] = None,
    ocpp_telemetry: Optional[Dict[str, Any]] = None,
    cms_billed_kwh: Optional[float] = None
) -> Dict[str, Any]:
    """
    Constructs an intelligent Root Cause Analysis diagnostic report with fault attribution,
    physical telemetry proof, and actionable corrective recommendations.
    Tariff is strictly taken from the Charging Details endpoint under the Per Unit Price section.
    """
    tariff_inr_per_kwh = float(tariff_inr_per_kwh or 0.0)
    b_kwh = float(cms_billed_kwh or total_energy_kwh)
    connector_kw = float(charger_specs.get("gun_capacity_kw") or charger_specs.get("rated_capacity_kw") or 30.0)
    rate_per_sec = connector_kw / 3600.0
    
    # Identify primary anomalies
    has_spike = any(a["type"] in ("ABNORMAL_ENERGY_JUMP", "EXCESS_ENERGY_SPIKE", "SPIKE_POWER_EXCEEDED") for a in anomalies)
    has_overrun = any(a["type"] == "BATTERY_CAPACITY_OVERRUN" for a in anomalies)
    has_rollback = any(a["type"] == "METER_ROLLBACK" for a in anomalies)
    has_inversion = any(a["type"] == "THERMODYNAMIC_INVERSION" for a in anomalies)
    has_zombie = any(a["type"] in ("ORPHANED_ZOMBIE_TRANSACTION", "ZOMBIE_SESSION_PHANTOM_OVERRUN") for a in anomalies) or (total_duration_sec > 14400 and total_energy_kwh < 0.05)
    has_meter_discrepancy = any(a["type"] in ("DISCREPANCY_BILLED_VS_METERED", "STOP_METER_LEAK_DISCREPANCY") for a in anomalies)
    has_zero = any(a["type"] in ("ZERO_ENERGY_DELIVERY", "ZERO_ENERGY_SESSION", "CONNECTOR_STATE_ABORT") for a in anomalies) or (total_energy_kwh < 0.05)

    if not anomalies and total_energy_kwh >= 0.05 and b_kwh >= 0.05:
        return {
            "status": "HEALTHY",
            "verdict": "NOMINAL: Verified Physical Parameters & Energy Curves",
            "confidence_score": 98,
            "badge_color": "green",
            "primary_cause": "Normal Operation",
            "summary_narrative": (
                f"Transaction #{transaction_id} demonstrated stable energy delivery conforming to physical connector transfer limits. "
                f"Total consumption of {total_energy_kwh:.2f} kWh was within the vehicle battery capacity ({vehicle_specs['capacity_kwh']:.1f} kWh), "
                f"and interval energy increments remained strictly compliant with the {connector_kw:.0f} kW connector rate "
                f"({rate_per_sec:.4f} kWh/s)."
            ),
            "evidence": [
                f"All interval energy increments <= physical {connector_kw:.0f} kW connector limit ({rate_per_sec:.4f} kWh/s)",
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

    evidence = []
    recommendations = []
    phantom_kwh = 0.0
    
    if has_spike:
        primary_cause = "Hardware / Modbus CT Ratio Surge or Erratic Meter Increment"
        spike_items = [a for a in anomalies if a["type"] in ("ABNORMAL_ENERGY_JUMP", "EXCESS_ENERGY_SPIKE", "SPIKE_POWER_EXCEEDED")]
        top_spike = max(spike_items, key=lambda x: x.get("delta_kwh", 0))
        delta_kwh = top_spike.get("delta_kwh", 0)
        dt_sec = top_spike.get("interval_sec", 0)
        expected_kwh = top_spike.get("expected_kwh", 0)
        ratio_pct = top_spike.get("ratio_pct", 0)
        
        evidence.append(
            f"Impossible energy jump of +{delta_kwh:.3f} kWh in {dt_sec:.0f}s recorded at {top_spike.get('timestamp')}, "
            f"exceeding physical {connector_kw:.0f} kW connector delivery ceiling ({expected_kwh:.3f} kWh) by {max(0, ratio_pct - 100):.0f}%."
        )
        evidence.append(
            f"Unitary connector limit: a {connector_kw:.0f} kW connector delivers at most {connector_kw:.0f} kWh in 1 hour "
            f"({(connector_kw/60):.2f} kWh/min, {rate_per_sec:.5f} kWh/s)."
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
        if tariff_inr_per_kwh > 0:
            recommendations.append(
                f"Credit customer account for {phantom_kwh:.2f} phantom kWh (approx ₹{phantom_kwh * tariff_inr_per_kwh:.2f} at Per Unit Price ₹{tariff_inr_per_kwh:.2f}/kWh from Charging Details endpoint)."
            )
        else:
            recommendations.append(
                f"Credit customer account for {phantom_kwh:.2f} phantom kWh based on the Per Unit Price from Charging Details endpoint."
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

    if has_zombie:
        primary_cause = "Orphaned Zombie Transaction & Phantom Meter Advance Accumulation"
        evidence.append(
            f"Physical meter delivered {total_energy_kwh:.3f} kWh to the vehicle across {intervals_count} reading(s), "
            f"yet CMS billed {b_kwh:.2f} kWh from Stop meter advancing during an orphaned session."
        )
        if total_duration_sec > 3600:
            evidence.append(
                f"Transaction hung open in CMS for {total_duration_sec/3600:.1f} hours ({total_duration_sec/86400:.1f} days) "
                f"without periodic MeterValues, accumulating other vehicles' charging consumption."
            )
        if ocpp_telemetry and ocpp_telemetry.get("status_transitions"):
            evidence.append("OCPP status telemetry indicates connector reverted to 'Available' within seconds of StartTransaction without power draw.")
        if has_overrun:
            overrun_item = next(a for a in anomalies if a["type"] == "BATTERY_CAPACITY_OVERRUN")
            phantom_kwh = max(phantom_kwh, overrun_item.get("phantom_kwh", b_kwh))
            evidence.append(
                f"Vehicle is {vehicle_specs['name']} with an {vehicle_specs['capacity_kwh']:.1f} kWh battery pack. "
                f"Billed {b_kwh:.2f} kWh represents a {((b_kwh/vehicle_specs['capacity_kwh'])-1)*100:.0f}% physical impossibility."
            )
        recommendations.append(
            "Waive / cancel pending CMS charge and issue full refund for any pre-auth deposit deducted from driver wallet."
        )
        recommendations.append(
            "Implement OCPP watchdog / auto-close rule to terminate active transactions if MeterValues telemetry ceases for > 15 minutes."
        )
        recommendations.append(
            "Verify charger firmware to ensure StopTransaction is transmitted upon connector release."
        )

    elif has_meter_discrepancy:
        primary_cause = "Billed vs. Metered Energy Discrepancy (Unmetered Stop Delta)"
        gap = max(0.0, b_kwh - total_energy_kwh)
        phantom_kwh = max(phantom_kwh, gap)
        evidence.append(
            f"Physical telemetry verified {total_energy_kwh:.3f} kWh delivered to the vehicle, "
            f"yet CMS billed {b_kwh:.2f} kWh (Difference: +{gap:.2f} kWh unmetered gap)."
        )
        if tariff_inr_per_kwh > 0:
            recommendations.append(
                f"Credit customer account for {gap:.2f} unmetered phantom kWh (approx ₹{gap * tariff_inr_per_kwh:.2f} at Per Unit Price ₹{tariff_inr_per_kwh:.2f}/kWh)."
            )
        else:
            recommendations.append(
                f"Credit customer account for {gap:.2f} unmetered phantom kWh based on the Per Unit Price from Charging Details endpoint."
            )
        recommendations.append(
            "Inspect charger controller firmware to ensure StopTransaction meterStop accurately reflects the final charging sample."
        )

    elif has_zero:
        primary_cause = "Zero-Energy Session Abort / Contactor Drop"
        evidence.append(f"Zero energy delivered (0.000 kWh) over {total_duration_sec/60:.1f} minutes of connected session.")
        recommendations.append("Inspect DC contactor coil drive signals and BMS authorization handshake.")
        recommendations.append("Ensure zero-energy aborted sessions are not billed to driver accounts.")

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
    
    # 1. Extract Per Unit Price strictly from Charging Details endpoint
    per_unit_price = 0.0
    tariff_source = "Charging Details Endpoint (Per Unit Price section)"
    
    if booking.get("PerUnitPrice") is not None:
        try:
            val = float(booking.get("PerUnitPrice"))
            if val > 0:
                per_unit_price = val
        except (ValueError, TypeError):
            pass

    if per_unit_price <= 0 and booking.get("UnitPrice") is not None:
        try:
            val = float(booking.get("UnitPrice"))
            if val > 0:
                per_unit_price = val
        except (ValueError, TypeError):
            pass

    # If PerUnitPrice not explicitly keyed, derive directly from Charging Details (ActualChargingAmount / EnergyConsumed)
    if per_unit_price <= 0:
        try:
            act_base = float(booking.get("ActualChargingAmount") or 0.0)
            act_kwh = float(booking.get("EnergyConsumed_indecimal") or booking.get("EnergyConsumed") or 0.0)
            if act_base > 0 and act_kwh > 0:
                per_unit_price = round(act_base / act_kwh, 2)
                tariff_source = "Charging Details Endpoint (Derived from ActualChargingAmount / EnergyConsumed)"
        except (ValueError, TypeError):
            pass

    if per_unit_price <= 0:
        tariff_source = "Charging Details Endpoint (Per Unit Price not specified / 0.00)" 

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

    pending_amt = 0.0
    if booking.get("PendingAmount") is not None:
        try:
            pending_amt = float(booking.get("PendingAmount") or 0.0)
        except (ValueError, TypeError):
            pass

    ref_kwh = cms_kwh if cms_kwh > 0.0 else total_energy_kwh
    
    # Zero delivery applies strictly when physical meter confirms 0 kWh delivered to the vehicle
    has_zero_delivery = (total_energy_kwh < 0.05) and (
        any(a.get("type") in ("ZERO_ENERGY_DELIVERY", "ZERO_ENERGY_SESSION", "CONNECTOR_STATE_ABORT") for a in analysis.get("anomalies", []))
        or total_energy_kwh == 0.0
    )

    if has_zero_delivery:
        # Zero physical energy delivered - real legitimate charge is strictly ₹0.00
        legitimate_kwh = 0.0
        real_base = 0.0
        real_tax = 0.0
        real_total_after_taxes = 0.0
    elif not is_anomaly and phantom_kwh <= 0.0:
        # 100% nominal session - physical conforms to billed
        legitimate_kwh = ref_kwh
        real_base = actual_base if actual_base > 0.0 else round(legitimate_kwh * per_unit_price, 2)
        real_tax = actual_tax if actual_tax > 0.0 else round(real_base * (tax_rate_pct / 100.0), 2)
        real_total_after_taxes = actual_user_paid if actual_user_paid > 0.0 else round(real_base + real_tax, 2)
    else:
        # Physical energy WAS delivered to the vehicle:
        # Ground legitimate energy in physical metered delivery, capped by battery capacity
        pack_cap = float(veh.get("capacity_kwh", 0.0))
        
        if total_energy_kwh >= 0.05:
            cand_kwh = min(total_energy_kwh, ref_kwh) if ref_kwh > 0 else total_energy_kwh
        else:
            cand_kwh = max(0.0, ref_kwh - phantom_kwh)

        # Cap by physical battery capacity if pack size is known
        if pack_cap > 0.0 and cand_kwh > pack_cap * 1.05:
            legitimate_kwh = pack_cap
        else:
            legitimate_kwh = cand_kwh

        real_base = round(legitimate_kwh * per_unit_price, 2)
        real_tax = round(real_base * (tax_rate_pct / 100.0), 2)
        real_total_after_taxes = round(real_base + real_tax, 2)

    variance = round(actual_user_paid - real_total_after_taxes, 2)
    base_discrepancy = round(actual_base - real_base, 2)
    tax_discrepancy = round(actual_tax - real_tax, 2)

    if variance > 0.50:
        status_code = "OVERCHARGED"
        badge_color = "red"
        if legitimate_kwh <= 0.0:
            status_label = f"CRITICAL OVERCHARGE (₹{variance:.2f} Discrepancy - 0.00 kWh Legitimate Energy)"
            action_summary = (
                f"Transaction delivered 0.000 kWh to the vehicle (Aborted / Zero-Energy session). "
                f"Real legitimate cost is ₹0.00 after all taxes. The user was billed ₹{actual_user_paid:.2f} "
                f"(Pending: ₹{pending_amt:.2f}, Paid: ₹{payment_rec:.2f}) due to an unclosed orphaned session accumulating "
                f"unmetered consumption. Immediate action: cancel/waive pending charge of ₹{pending_amt:.2f} and refund ₹{payment_rec:.2f} to user wallet."
            )
        else:
            phantom_units = round(max(0.0, ref_kwh - legitimate_kwh), 3)
            status_label = f"OVERCHARGED BY ₹{variance:.2f} (+{phantom_units:.2f} Phantom kWh)"
            action_summary = (
                f"User was billed for {ref_kwh:.2f} kWh (₹{actual_user_paid:.2f}), but physical telemetry verifies that only "
                f"{legitimate_kwh:.2f} kWh was legitimately delivered into the vehicle. "
                f"The customer was overcharged by ₹{variance:.2f} (Base: ₹{base_discrepancy:.2f} + Tax: ₹{tax_discrepancy:.2f}) "
                f"for {phantom_units:.2f} phantom/unmetered kWh. A commercial credit/refund of ₹{variance:.2f} is due to the user."
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
        "tariff_source": tariff_source,
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
            "pending_amount_inr": round(pending_amt, 2),
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

    # 3. Extract Per Unit Price from Charging Details endpoint
    tariff_from_charging_details = None
    if booking_record:
        if booking_record.get("PerUnitPrice") is not None:
            try:
                tariff_from_charging_details = float(booking_record.get("PerUnitPrice"))
            except (ValueError, TypeError):
                pass
        if (tariff_from_charging_details is None or tariff_from_charging_details <= 0) and booking_record.get("UnitPrice") is not None:
            try:
                tariff_from_charging_details = float(booking_record.get("UnitPrice"))
            except (ValueError, TypeError):
                pass
        if (tariff_from_charging_details is None or tariff_from_charging_details <= 0) and booking_record.get("ActualChargingAmount") and (booking_record.get("EnergyConsumed_indecimal") or booking_record.get("EnergyConsumed")):
            try:
                c_amt = float(booking_record.get("ActualChargingAmount"))
                c_kwh = float(booking_record.get("EnergyConsumed_indecimal") or booking_record.get("EnergyConsumed"))
                if c_amt > 0 and c_kwh > 0:
                    tariff_from_charging_details = round(c_amt / c_kwh, 2)
            except (ValueError, TypeError):
                pass

    # Fetch and correlate full OCPP telemetry & meter values payloads
    ocpp_telemetry = fetch_and_correlate_ocpp_telemetry(
        party_id=party_id,
        charger_id=charger_id,
        charger_code=charger_code,
        tx_id=tx_id,
        start_time=start_time,
        stop_time=stop_time,
        booking_record=booking_record,
        tx_record=tx_record
    )

    # Perform energy anomaly detection with exact Charging Details Per Unit Price & OCPP correlation
    analysis = analyze_transaction_energy(
        party_id=party_id,
        transaction_id=tx_id,
        charger_code=charger_code,
        charger_id=charger_id,
        start_time=start_time,
        stop_time=stop_time,
        vehicle_model_hint=vehicle_str,
        custom_battery_kwh=battery_kwh,
        tx_meta=tx_record,
        tariff_per_unit=tariff_from_charging_details,
        booking_record=booking_record,
        ocpp_telemetry=ocpp_telemetry
    )

    analysis["ocpp_telemetry"] = ocpp_telemetry

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
