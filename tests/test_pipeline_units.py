import re
import json
import time
import threading

class HostRateLimiter:
    def __init__(self, min_interval_seconds: float = 0.20):
        self.min_interval = min_interval_seconds
        self.last_request_time: dict[str, float] = {}
        self.lock = threading.Lock()

    def acquire(self, host: str):
        with self.lock:
            now = time.time()
            last = self.last_request_time.get(host, 0.0)
            elapsed = now - last
            if elapsed < self.min_interval:
                sleep_time = self.min_interval - elapsed
                time.sleep(sleep_time)
            self.last_request_time[host] = time.time()

def classify_ocpi_text(text: str) -> tuple[bool, str]:
    clean = str(text or "").strip()
    if not clean or clean.lower() in ("null", "none", "n/a", "na", "-", "undefined", "noerror", "ok", "success"):
        return False, ""
    # Check for UUID
    if re.match(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$", clean):
        return False, clean
    # Check for pure hex/alphanumeric token without spaces
    if len(clean) >= 12 and re.match(r"^[0-9a-zA-Z_-]+$", clean) and not any(k in clean.lower() for k in ["error", "fail", "reject", "busy", "timeout", "offline", "abort", "fault"]):
        return False, clean
    lower = clean.lower()
    known = [
        "not connected", "disconnected", "busy", "occupied", "rejected", "timeout", "timed out",
        "invalid", "blocked", "unauthorized", "offline", "unreachable", "abort", "fault",
        "error", "failure", "fail", "stopped", "expired", "limit", "denied", "unknown", "refused",
        "start initiated"
    ]
    if any(k in lower for k in known):
        # "start initiated" is status info, not failure
        if "start initiated" in lower:
            return False, clean
        return True, clean
    words = clean.split()
    if len(words) >= 2 and all(w.isalnum() for w in words):
        return True, clean
    return False, clean

def parse_ocpi_start_session(raw_text: str) -> dict:
    text = str(raw_text or "").strip()
    clean_raw = text.split("??")[0].strip()
    if not clean_raw or clean_raw.upper() in ("N/A", "NA", "NULL", "NONE"):
        return {
            "is_na": True,
            "raw_response": text,
            "result": "NA",
            "ocpi_text": "",
            "is_real_reason": False,
            "failure_reason": ""
        }
    try:
        data = json.loads(clean_raw)
        item = data[0] if isinstance(data, list) and data else (data if isinstance(data, dict) else {})
        res_status = str(item.get("result") or item.get("status") or "").upper()
        messages = item.get("message") or []
        texts = []
        if isinstance(messages, list):
            for m in messages:
                if isinstance(m, dict) and m.get("text"):
                    texts.append(str(m["text"]).strip())
                elif isinstance(m, str):
                    texts.append(m.strip())
        combined_text = " | ".join(texts)
        is_reason, failure_text = classify_ocpi_text(combined_text)
        return {
            "is_na": False,
            "raw_response": clean_raw,
            "result": res_status,
            "ocpi_text": combined_text,
            "is_real_reason": is_reason,
            "failure_reason": failure_text if is_reason else ""
        }
    except Exception:
        # Check if text contains NA
        if "N/A" in clean_raw or "NA" in clean_raw:
            return {
                "is_na": True,
                "raw_response": text,
                "result": "NA",
                "ocpi_text": "",
                "is_real_reason": False,
                "failure_reason": ""
            }
        is_reason, failure_text = classify_ocpi_text(clean_raw)
        return {
            "is_na": False,
            "raw_response": clean_raw,
            "result": "UNKNOWN",
            "ocpi_text": clean_raw,
            "is_real_reason": is_reason,
            "failure_reason": failure_text if is_reason else ""
        }


def test_ocpi_start_session_na():
    assert parse_ocpi_start_session('N/A ?? "N/A"')["is_na"] is True
    assert parse_ocpi_start_session('NA')["is_na"] is True
    assert parse_ocpi_start_session('')["is_na"] is True

def test_ocpi_start_session_uuid():
    uuid_sample = '[{"result":"ACCEPTED","timeout":60,"message":[{"language":"en","text":"23b050ba-bebe-4c85-84f1-2fdffb7b02c3"}],"BookingId":0}] ?? "N/A"'
    p_uuid = parse_ocpi_start_session(uuid_sample)
    assert p_uuid["is_na"] is False
    assert p_uuid["result"] == "ACCEPTED"
    assert p_uuid["is_real_reason"] is False

def test_ocpi_start_session_error():
    error_sample = '[{"result":"REJECTED","timeout":60,"message":[{"language":"en","text":"EV not connected"}],"BookingId":0}] ?? "N/A"'
    p_err = parse_ocpi_start_session(error_sample)
    assert p_err["is_na"] is False
    assert p_err["result"] == "REJECTED"
    assert p_err["is_real_reason"] is True
    assert p_err["failure_reason"] == "EV not connected"

def test_filter_completed_low_consumption():
    def filter_completed_row(kwh: float, is_completed_sheet: bool):
        if is_completed_sheet and kwh >= 1.0:
            return False, "Discarded (Normal Session >= 1.0 kWh)"
        if is_completed_sheet:
            return True, "Completed (<1kWh)"
        return True, "Cancelled"

    assert filter_completed_row(18.5, True)[0] is False
    assert filter_completed_row(1.0, True)[0] is False
    assert filter_completed_row(0.99, True)[0] is True
    assert filter_completed_row(0.35, True)[1] == "Completed (<1kWh)"


def test_vehicle_model_aggregation_disambiguation():
    from src.rca.roaming_upload_analyzer import RoamingUploadAnalyzer
    analyzer = RoamingUploadAnalyzer()
    instances = [
        {
            "booking_id": "1001",
            "resolved_evse_id": "EVSE-01",
            "roaming_uid": "IOC-01",
            "party_id": "IOC",
            "station_name": "Delhi Central",
            "vehicle": "Nexon EV (DL01AB1234)",
            "vehicle_make": "Tata Motors",
            "vehicle_model": "Nexon EV",
            "manufacturer": "Tata Motors",
            "charger_model": "Delta UFC50",
            "charger_manufacturer": "DELTA",
            "root_cause": "BMS Handshake Timeout (~60s Protocol Window)",
            "fault_side": "VEHICLE SIDE",
            "kwh": 0.0,
            "session_category": "Cancelled"
        },
        {
            "booking_id": "1002",
            "resolved_evse_id": "EVSE-02",
            "roaming_uid": "IOC-02",
            "party_id": "IOC",
            "station_name": "Noida Hub",
            "vehicle": "XEV 9e (UP16XY9999)",
            "vehicle_make": "Mahindra",
            "vehicle_model": "XEV 9e",
            "manufacturer": "Mahindra",
            "charger_model": "Exicom Harmony",
            "charger_manufacturer": "EXICOM",
            "root_cause": "Charger Controller Initiation Timeout",
            "fault_side": "CHARGER SIDE",
            "kwh": 0.0,
            "session_category": "Cancelled"
        }
    ]

    aggs = analyzer.compute_aggregations(instances)
    v_stats = aggs["vehicle_model_stats"]
    assert len(v_stats) == 2
    tata = next(it for it in v_stats if it["mfg"] == "Tata Motors")
    assert tata["model"] == "Nexon EV"
    assert tata["vehicle_bms_faults"] == 1
    assert tata["dominant_side"] == "Vehicle Side"

    mahindra = next(it for it in v_stats if it["mfg"] == "Mahindra")
    assert mahindra["model"] == "XEV 9e"
    assert mahindra["charger_side"] == 1
    assert mahindra["vehicle_bms_faults"] == 0

    hw_stats = aggs["charger_hardware_stats"]
    assert len(hw_stats) == 2
    delta = next(it for it in hw_stats if it["mfg"] == "DELTA")
    assert delta["model"] == "Delta UFC50"
    exicom = next(it for it in hw_stats if it["mfg"] == "EXICOM")
    assert exicom["model"] == "Exicom Harmony"
    assert exicom["charger_hardware_faults"] == 1


def test_resolve_charger_hardware_lookup():
    from src.rca.roaming_upload_analyzer import RoamingUploadAnalyzer
    analyzer = RoamingUploadAnalyzer()

    # Test brand signature resolution
    res_tirex = analyzer.resolve_charger_hardware(station_name="IOCL Tirex Speedway 240kW")
    assert res_tirex["charger_manufacturer"] == "Tirex"
    assert "240" in res_tirex["charger_model"]

    res_exicom = analyzer.resolve_charger_hardware(station_name="Adhoc Exicom DC Fast Charging")
    assert res_exicom["charger_manufacturer"] == "EXICOM"

    res_delta = analyzer.resolve_charger_hardware(station_name="Highway Delta CCS240 Plaza")
    assert res_delta["charger_manufacturer"] == "DELTA"


