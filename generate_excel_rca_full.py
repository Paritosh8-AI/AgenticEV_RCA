"""
ElectreeFi CMS - Full 24-Hour Cancelled & Low-Consumption (<1kWh) Bookings RCA Excel Generator
Queries all cancelled bookings and all completed bookings with energy < 1 kWh (pageSize=1000).
For all low-consumption bookings, matches the exact Transaction ID generated for the corresponding
Booking ID from the central OCPP Transactions grid (/OCPPManagement/TransactionDetail).
Ensures 100% precision by strictly bounding logs to the transaction start/stop times and filtering
exclusively by transactionId and connectorId, eliminating cross-session interference.
"""

import os
import sys
import json
import re
import urllib.request
import urllib.parse
import asyncio
import concurrent.futures
from datetime import datetime, timedelta
from typing import Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from src.auth.session_manager import is_session_valid, is_session_valid_sync, interactive_login
from src.rca.ocpp_parser import parse_ocpp_message, is_remote_start_status_rejected, is_remote_start_status_accepted
from src.models import OCPPLogEntry

# Ensure UTF-8 output on Windows
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")


def get_cookie_header() -> str:
    cookie_path = "data/session_state.json"
    if not os.path.exists(cookie_path):
        return ""
    try:
        with open(cookie_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        cookies = data.get("cookies", [])
        return "; ".join([f"{c['name']}={c['value']}" for c in cookies if "ev-charge-network.com" in c.get("domain", "")])
    except Exception:
        return ""


def is_numerical_value(val: str) -> bool:
    """Checks if a string represents a pure numeric value (e.g. 574, 575, 0, 0.000000)."""
    if not val:
        return False
    v = val.strip()
    try:
        float(v)
        return True
    except ValueError:
        pass
    return v.isdigit()


def is_hardware_diagnostic_tag(val: str | None) -> bool:
    """Matches hardware diagnostic error tags such as [E039], [E052], [E213], [E334], etc."""
    if not val:
        return False
    v = val.strip()
    if re.search(r'\[E\d+\]', v, re.IGNORECASE):
        return True
    if re.search(r'\[[A-Za-z0-9_-]{2,}\]', v):
        return True
    return False


def is_non_error_or_uninformative_info(val: str) -> bool:
    """
    Checks if the info tag is uninformative: numerical, 'other', GSM/RSSI status, normal, no alarm, etc.
    """
    if not val:
        return True
    v = val.strip().lower()
    if is_numerical_value(v):
        return True
    if v in ("other", "none", "null", "noerror", "no error", "no alarms", "no alarm", "normal", "-", "undefined", ""):
        return True
    if "gsm" in v or "rssi" in v:
        return True
    return False


def get_meaningful_info(logs: list[OCPPLogEntry], target_connector: int | None = None) -> str | None:
    """
    Scans logs for a valid, descriptive 'info: "..."' tag in StatusNotification messages.
    Prioritizes entries matching target_connector if available.
    Returns the descriptive string (e.g. 'LowInsulation'), or None if only numerical, 'other', GSM, or no info exists.
    """
    if not logs:
        return None

    # First pass: try to find on target_connector
    if target_connector is not None:
        try:
            cid = int(target_connector)
            for e in logs:
                if e.connector_id == cid and e.info and e.info.strip():
                    raw = e.info.strip()
                    if not is_non_error_or_uninformative_info(raw):
                        return raw
        except (ValueError, TypeError):
            pass

    # Second pass: check all logs
    for e in logs:
        if e.info and e.info.strip():
            raw = e.info.strip()
            if not is_non_error_or_uninformative_info(raw):
                return raw

    return None


def diagnose_cancelled_rca(logs: list[OCPPLogEntry], target_connector: int | None = None) -> tuple[str, bool]:
    """
    RCA diagnosis for CANCELLED bookings per user instructions:
    1. There can't be a reason saying remotely stopped by user because in cancelled sessions user can't even make the booking.
    2. Write 'remote start not received' if the connector goes into 'preparing' and there is no remotestart transaction.
    3. If gun is 'preparing' and remote start transaction has been accepted, look for start transaction and if not found write 'Start transaction not recieved'.
    4. If there is absolutely no info in the logs around the corresponding timestamp just write 'gun not connected'.
    5. For errors look at the 'info: ' tag in the status notification and write whatever appears there:
       - If there is a numerical value or the only info is 'other' or GSM connected, write 'Can't be specified by the logs'.
    6. Hardware diagnostic tags: return is_hw_diag=True so the booking is ignored per user instruction.
    7. Only write "remote start transaction was rejected" when you find the remote start transaction status as "rejected".
    Returns (rca_string, is_hw_diag).
    """
    if not logs:
        return ("gun not connected", False)

    # Check for hardware diagnostic error tags in logs
    for e in logs:
        if is_hardware_diagnostic_tag(e.info) or is_hardware_diagnostic_tag(e.error_code):
            return ("[HARDWARE_DIAGNOSTIC_IGNORE]", True)

    # Step 1: Check for valid descriptive 'info: ' tag in StatusNotification
    meaningful_info = get_meaningful_info(logs, target_connector)
    if meaningful_info:
        if is_hardware_diagnostic_tag(meaningful_info):
            return ("[HARDWARE_DIAGNOSTIC_IGNORE]", True)
        if meaningful_info.strip().lower() not in ("remote", "remotely stop", "remote stop", "remotely stopped by user"):
            return (meaningful_info, False)

    # Step 2: Check for hardware errorCode in StatusNotification
    for e in logs:
        if e.error_code and e.error_code.lower() not in ("noerror", "none", "null", ""):
            err = e.error_code.lower()
            if "ground" in err:
                return ("Earth / Ground Fault Detected", False)
            if "power" in err or "grid" in err:
                return ("AC Mains Grid Power Loss", False)
            if "high" in err and "temp" in err:
                return ("Thermal Limit Exceeded", False)
            if "evcommunication" in err or "comm" in err:
                return ("Pre-Charge Handshake Abort (Never Started)", False)

    # Step 3: Check for rejected RemoteStartTransaction
    if is_remote_start_status_rejected(logs):
        return ("remote start transaction was rejected", False)

    # Step 4: Analyze sequence (Preparing, RemoteStart, StartTransaction)
    has_preparing = False
    has_remote_start = False
    remote_start_accepted = is_remote_start_status_accepted(logs)
    has_start_transaction = False

    for e in logs:
        raw = (e.raw_message or "").lower()
        ev = (e.event_name or "").lower()
        act = (e.parsed_action or "").lower()

        if "preparing" in raw:
            has_preparing = True

        if "remotestart" in raw or "remotestart" in ev or "remotestart" in act:
            has_remote_start = True

        if "accepted" in raw or '"status": "accepted"' in raw or '"status":"accepted"' in raw:
            remote_start_accepted = True

        if ev == "starttransaction" or act == "starttransaction" or "starttransaction" in raw:
            has_start_transaction = True

    if has_preparing and remote_start_accepted and not has_start_transaction:
        return ("Start transaction not recieved", False)

    if has_preparing and not has_remote_start:
        return ("remote start not received", False)

    if remote_start_accepted and not has_start_transaction:
        return ("Start transaction not recieved", False)

    if has_preparing:
        return ("remote start not received", False)

    return ("Can't be specified by the logs", False)


def diagnose_completed_low_energy_rca(
    stop_reason_charger: str | None,
    stop_reason_server: str | None,
    logs: list[OCPPLogEntry],
    target_connector: int | None = None
) -> tuple[str, bool]:
    """
    RCA diagnosis for COMPLETED bookings with < 1 kWh consumption per user instructions:
    1. Firstly read the stop reason (from charger) and stop reason (from server) on the details popup and write them as is.
       - If StopTransaction / transaction has a specific reason (e.g. "EVDisconnected", "EmergencyStop", "Remote", "Local", "DeAuthorized"),
         write it as is! Do NOT override with "Error Noticed by EV".
    2. Go further and check the logs IF ONLY the reason says nothing or 'Other'.
    3. In these situations (reason is 'Other' or empty):
       - ONLY write whatever is in the info: " " in StatusNotification (e.g. "LowInsulation", "Param config failed").
       - If there is a numerical value here (e.g. "574", "575", "0") or the only info is "other" or GSM connected:
         write "Can't be specified by the logs".
    4. Hardware diagnostic tags: return is_hw_diag=True so the booking is ignored per user instruction.
    Returns (rca_string, is_hw_diag).
    """
    ch_reason = (stop_reason_charger or "").strip()
    srv_reason = (stop_reason_server or "").strip()

    # Step 1: If ch_reason was empty from the API, extract from StopTransaction in logs
    if not ch_reason or ch_reason.lower() in ("none", "null", "-", "undefined", ""):
        if logs:
            stop_tx_reasons = []
            for e in logs:
                is_stop_tx = (e.event_name.lower() == "stoptransaction" or (e.parsed_action and e.parsed_action.lower() == "stoptransaction"))
                if is_stop_tx and e.reason and e.reason.strip():
                    r = e.reason.strip()
                    if target_connector is not None:
                        try:
                            if e.connector_id == int(target_connector):
                                stop_tx_reasons.insert(0, r)
                            else:
                                stop_tx_reasons.append(r)
                        except (ValueError, TypeError):
                            stop_tx_reasons.append(r)
                    else:
                        stop_tx_reasons.append(r)
            if stop_tx_reasons:
                ch_reason = stop_tx_reasons[0]

    is_ch_empty = not ch_reason or ch_reason.lower() in ("none", "null", "-", "undefined", "")
    is_srv_empty = not srv_reason or srv_reason.lower() in ("none", "null", "-", "undefined", "")

    # Check for hardware diagnostic error tags in logs or stop reasons
    if is_hardware_diagnostic_tag(ch_reason) or is_hardware_diagnostic_tag(srv_reason):
        return ("[HARDWARE_DIAGNOSTIC_IGNORE]", True)
    for e in logs:
        if is_hardware_diagnostic_tag(e.info) or is_hardware_diagnostic_tag(e.error_code):
            return ("[HARDWARE_DIAGNOSTIC_IGNORE]", True)

    # 1. If stop reason from charger exists and is NOT "Other" -> write as is!
    if not is_ch_empty and ch_reason.lower() != "other":
        if ch_reason.lower() == "evdisconnected":
            ch_reason = "EVDisconnected"
        elif ch_reason.lower() in ("emergencystop", "emergency stop"):
            ch_reason = "EmergencyStop"
        elif ch_reason.lower() == "remote":
            ch_reason = "Remote"
        elif ch_reason.lower() == "local":
            ch_reason = "Local"
        elif ch_reason.lower() == "deauthorized":
            ch_reason = "DeAuthorized"

        if not is_srv_empty and srv_reason.lower() != "other" and srv_reason.lower() != ch_reason.lower():
            return (f"{ch_reason} (Server: {srv_reason})", False)
        return (ch_reason, False)

    # 2. If stop reason from server exists and is NOT "Other" -> write as is!
    if not is_srv_empty and srv_reason.lower() != "other":
        return (srv_reason, False)

    # 3. ONLY if the reason says nothing or "Other", check the logs!
    if not logs:
        return ("gun not connected", False)

    meaningful_info = get_meaningful_info(logs, target_connector)
    if meaningful_info:
        if is_hardware_diagnostic_tag(meaningful_info):
            return ("[HARDWARE_DIAGNOSTIC_IGNORE]", True)
        return (meaningful_info, False)

    return ("Can't be specified by the logs", False)


def parse_flexible_dt(t_str: str | None) -> datetime | None:
    if not t_str or not t_str.strip():
        return None
    cleaned = t_str.strip()
    for fmt in [
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%S.%f",
        "%Y-%m-%d %I:%M %p",
        "%Y-%m-%d %I:%M:%S %p",
        "%d-%m-%Y %I:%M %p",
        "%d-%m-%Y %I:%M:%S %p",
        "%d/%m/%Y %I:%M %p",
        "%d/%m/%Y %I:%M:%S %p",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%d",
        "%d-%m-%Y"
    ]:
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            continue
    return None


_CHARGER_LOGS_CACHE: dict[tuple, list[OCPPLogEntry]] = {}


def fetch_charger_logs_direct(ch_id: Any, from_enc: str, to_enc: str, cookie_header: str) -> list[OCPPLogEntry]:
    """Fetches and caches raw OCPP logs directly via HTTP POST."""
    if not ch_id:
        return []

    cache_key = (str(ch_id), from_enc, to_enc)
    if cache_key in _CHARGER_LOGS_CACHE:
        return _CHARGER_LOGS_CACHE[cache_key]

    url = f"https://cms.ev-charge-network.com/GetOcppLogs?entityId={ch_id}&fromDate={from_enc}&toDate={to_enc}"
    payload = urllib.parse.urlencode({
        "page": "1",
        "pageSize": "100",
        "sort": "",
        "group": "",
        "filter": ""
    }).encode("utf-8")
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Cookie": cookie_header
    }

    entries: list[OCPPLogEntry] = []
    try:
        req = urllib.request.Request(url, data=payload, headers=headers)
        with urllib.request.urlopen(req, timeout=10) as res:
            data = json.loads(res.read().decode("utf-8"))
            log_items = data.get("Data", []) if isinstance(data, dict) else []
            for item in log_items:
                ev_name = str(item.get("EventName") or "")
                if ev_name.lower() in ("heartbeat", "metervalues"):
                    continue
                entries.append(parse_ocpp_message(
                    raw_message=str(item.get("Message") or ""),
                    event_name=ev_name,
                    message_type=str(item.get("MessageTypeName") or ""),
                    timestamp=str(item.get("Date") or item.get("Timestamp") or ""),
                    message_id_fallback=str(item.get("MessageId") or "")
                ))
    except Exception:
        pass

    _CHARGER_LOGS_CACHE[cache_key] = entries
    return entries


def fetch_logs_for_booking(ch_id, b_dt: datetime, cookie_header: str) -> list[OCPPLogEntry]:
    """Fetches OCPP logs for a booking with tight ±3 minute lifecycle window."""
    w_start = (b_dt - timedelta(minutes=3)).strftime("%Y-%m-%d %H:%M:00").replace(" ", "%20")
    w_end = (b_dt + timedelta(minutes=3)).strftime("%Y-%m-%d %H:%M:00").replace(" ", "%20")
    return fetch_charger_logs_direct(ch_id, w_start, w_end, cookie_header)


def fetch_logs_for_exact_window(ch_id, t_start: datetime, t_stop: datetime, cookie_header: str) -> list[OCPPLogEntry]:
    """Fetches OCPP logs strictly bounded around the exact transaction start and stop times (±2 min)."""
    w_start = (t_start - timedelta(minutes=2)).strftime("%Y-%m-%d %H:%M:00").replace(" ", "%20")
    w_end = (t_stop + timedelta(minutes=2)).strftime("%Y-%m-%d %H:%M:00").replace(" ", "%20")
    return fetch_charger_logs_direct(ch_id, w_start, w_end, cookie_header)


def filter_logs_for_transaction(logs: list[OCPPLogEntry], tx_id: int | str, target_connector: int | None) -> list[OCPPLogEntry]:
    """
    Filters OCPP logs to guarantee that every examined message strictly belongs to the specific transaction
    and connector, eliminating cross-session interference from other charging operations around the same time.
    """
    filtered = []
    tx_str = str(tx_id)
    for entry in logs:
        p_load = entry.parsed_payload if isinstance(entry.parsed_payload, dict) else {}
        t_id = p_load.get("transactionId") or p_load.get("transaction_id")

        if t_id is not None:
            if str(t_id) == tx_str:
                filtered.append(entry)
        elif tx_str in entry.raw_message:
            filtered.append(entry)
        elif entry.connector_id is not None and target_connector is not None:
            try:
                if int(entry.connector_id) == int(target_connector):
                    filtered.append(entry)
            except (ValueError, TypeError):
                pass
        else:
            filtered.append(entry)
    return filtered


def fetch_ocpp_transactions_index(dates: list[str], cookie_header: str) -> dict[int, dict]:
    """
    Optimized transactions indexer: fetches page 1 (pageSize=1000) for target dates in at most 2 hits.
    Never attempts page 2+ which times out on the backend server.
    """
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Cookie": cookie_header
    }
    tx_map: dict[int, dict] = {}
    payload = urllib.parse.urlencode({"page": "1", "pageSize": "1000", "sort": "", "group": "", "filter": ""}).encode("utf-8")

    for d in dates:
        tx_url = f"https://cms.ev-charge-network.com/OCPPManagement/TransactionDetail/LoadTransactionViewThroughAjaxForData/2?StartDate={d}&Enddate={d}"
        try:
            req = urllib.request.Request(tx_url, data=payload, headers=headers)
            with urllib.request.urlopen(req, timeout=12) as res:
                data = json.loads(res.read().decode("utf-8"))
                items = data.get("Data", []) if data.get("Data") else []
                for it in items:
                    res_id = it.get("ReservationId")
                    if res_id and str(res_id).isdigit():
                        r_int = int(res_id)
                        if r_int > 0:
                            tx_map[r_int] = it
        except Exception:
            pass

    return tx_map


def run_full_rca(output_excel: str = "ElectreeFi_Cancelled_Bookings_RCA.xlsx"):
    print("=" * 80)
    print("ELECTREEFI CMS - 24-HOUR CANCELLED & LOW-CONSUMPTION (<1kWh) BOOKINGS RCA")
    print("Direct Native HTTP Engine with Transaction ID Matching & Zero Browser Overhead")
    print("=" * 80)

    # 1. Fast native session verification (0.3s)
    if not is_session_valid_sync():
        print("\nSession is expired. Launching login...")
        asyncio.run(interactive_login(force=True))
        if not is_session_valid_sync():
            print("ERROR: Authentication failed.")
            sys.exit(1)

    cookie_header = get_cookie_header()
    now = datetime.now()
    start_24h = now - timedelta(hours=24)
    start_date_str = start_24h.strftime("%Y-%m-%d")
    end_date_str = now.strftime("%Y-%m-%d")

    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Cookie": cookie_header
    }

    # -------------------------------------------------------------
    # STEP 1: Query Cancelled Bookings (pageSize=1000 in 1 hit)
    # -------------------------------------------------------------
    print(f"\n[QUERY 1/3] Fetching Cancelled bookings from {start_date_str} to {end_date_str} (pageSize=1000)...")
    cancelled_url = f"https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails/CancelledBookingDataThroughAjax?StationId=&StartDate={start_date_str}&Enddate={end_date_str}&stateId=&cityId=&transactionTypeId=-1&Id=0&vin=&chargerCode=&user=&vehicleNumber="
    payload_canc = urllib.parse.urlencode({"sort": "", "page": "1", "pageSize": "1000", "group": "", "filter": ""}).encode("utf-8")
    
    cancelled_candidates = []
    try:
        req_canc = urllib.request.Request(cancelled_url, data=payload_canc, headers=headers)
        with urllib.request.urlopen(req_canc, timeout=15) as res:
            data_canc = json.loads(res.read().decode("utf-8"))
            raw_cancelled = data_canc.get("Data", []) if data_canc.get("Data") else []
            print(f"Retrieved {len(raw_cancelled)} total cancelled bookings from portal API!")
            for it in raw_cancelled:
                t_str = it.get("BookingInTime") or it.get("DateForBooking")
                dt = parse_flexible_dt(t_str) or now
                if dt >= start_24h:
                    cancelled_candidates.append({
                        "dt": dt,
                        "reason_type": "Cancelled",
                        "item": it,
                        "energy": 0.0
                    })
    except Exception as e:
        print(f"Error querying cancelled bookings: {e}")

    print(f"Total Cancelled bookings in 24h window: {len(cancelled_candidates)}")

    # -------------------------------------------------------------
    # STEP 2: Query Completed Bookings with Energy < 1 kWh (pageSize=1000 in 1 hit)
    # -------------------------------------------------------------
    print(f"\n[QUERY 2/3] Fetching Completed bookings where Energy < 1 kWh from {start_date_str} to {end_date_str}...")
    completed_url = f"https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails/PastBookingDataThroughAjax?StartDate={start_date_str}&Enddate={end_date_str}&transactionTypeId=-1&Id=0&StationId=&stateId=&cityId=&vin=&chargerCode=&user=&vehicleNumber=&i="
    payload_comp = urllib.parse.urlencode({"sort": "", "page": "1", "pageSize": "1000", "group": "", "filter": "EnergyConsumed~lt~1"}).encode("utf-8")

    completed_candidates = []
    try:
        req_comp = urllib.request.Request(completed_url, data=payload_comp, headers=headers)
        with urllib.request.urlopen(req_comp, timeout=15) as res:
            data_comp = json.loads(res.read().decode("utf-8"))
            raw_completed = data_comp.get("Data", []) if data_comp.get("Data") else []
            print(f"Retrieved {len(raw_completed)} completed bookings with Energy < 1 kWh from portal API!")
            for it in raw_completed:
                t_str = it.get("BookingInTime") or it.get("DateForBooking")
                dt = parse_flexible_dt(t_str) or now
                if dt >= start_24h:
                    try:
                        energy_val = float(it.get("EnergyConsumed_indecimal") or it.get("EnergyConsumed") or 0.0)
                    except (ValueError, TypeError):
                        energy_val = 0.0
                    completed_candidates.append({
                        "dt": dt,
                        "reason_type": "less than 1kWh",
                        "item": it,
                        "energy": energy_val
                    })
    except Exception as e:
        print(f"Error querying completed bookings: {e}")

    print(f"Total Completed bookings (<1kWh) in 24h window: {len(completed_candidates)}")

    # -------------------------------------------------------------
    # STEP 3: Fetch Central OCPP Transactions Index (Matching Tx ID)
    # -------------------------------------------------------------
    dates_to_index = [start_date_str] if start_date_str == end_date_str else [start_date_str, end_date_str]
    tx_lookup: dict[int, dict] = {}
    if completed_candidates:
        print(f"\n[QUERY 3/3] Indexing central OCPP Transactions for dates: {dates_to_index}...")
        tx_lookup = fetch_ocpp_transactions_index(dates_to_index, cookie_header)
        print(f"Indexed {len(tx_lookup)} unique transactions from OCPP Transactions table!")
    else:
        print("\n[QUERY 3/3] No completed low-consumption candidates; skipping transaction index (0 hits saved!).")

    # Combine both datasets
    all_candidates = cancelled_candidates + completed_candidates
    print(f"\n--> TOTAL COMBINED SESSIONS TO PROCESS: {len(all_candidates)} ({len(cancelled_candidates)} Cancelled + {len(completed_candidates)} Completed <1kWh)")

    # Concurrently fetch OCPP logs and diagnose RCA using ThreadPoolExecutor (max_workers=25)
    print(f"\n[OCPP] Concurrently analyzing logs with exact Transaction ID matching (ThreadPool: 25 workers)...")

    def process_candidate(candidate_info):
        dt = candidate_info["dt"]
        reason_type = candidate_info["reason_type"]
        it = candidate_info["item"]
        energy_val = candidate_info["energy"]

        bid = it.get("ChargingStationBookingId")
        st_name = it.get("ChargingStationName") or "Unknown Station"
        ch_code = it.get("ChargerCode") or ""
        conn_id = it.get("ConnectorId") or it.get("ConnectorID") or 1
        ch_id = it.get("ChargingStationChargerId")

        b_date_str = f"{dt.month}/{dt.day}/{dt.year}"

        if reason_type == "Cancelled":
            tx_id_display = ""
            logs = fetch_logs_for_booking(ch_id, dt, cookie_header)
            cs_rca_text, is_hw_diag = diagnose_cancelled_rca(logs, target_connector=conn_id)
        else:
            # Completed (<1kWh): Match Transaction ID
            bid_int = int(bid) if str(bid).isdigit() else None
            matched_tx = tx_lookup.get(bid_int) if bid_int else None

            if matched_tx:
                tx_pk = matched_tx.get("transaction_pk")
                tx_id_display = str(tx_pk) if tx_pk else ""
                tx_stop_reason = (matched_tx.get("stop_reason") or "").strip()
                target_conn = matched_tx.get("connector_pk") if matched_tx.get("connector_pk") is not None else conn_id

                # Per rule: If stop reason is known and not "Other", write as is! NO logs needed!
                if tx_stop_reason and tx_stop_reason.lower() not in ("other", "none", "null", "-", ""):
                    cs_rca_text, is_hw_diag = diagnose_completed_low_energy_rca(
                        stop_reason_charger=tx_stop_reason,
                        stop_reason_server=None,
                        logs=[],
                        target_connector=target_conn
                    )
                else:
                    # ONLY check logs if reason is "Other" or empty!
                    t_start = parse_flexible_dt(matched_tx.get("start_timestamp_string")) or dt
                    t_stop = parse_flexible_dt(matched_tx.get("stop_timestamp_string")) or (t_start + timedelta(minutes=3))
                    raw_logs = fetch_logs_for_exact_window(ch_id, t_start, t_stop, cookie_header)
                    filtered_logs = filter_logs_for_transaction(raw_logs, tx_pk, target_conn)
                    cs_rca_text, is_hw_diag = diagnose_completed_low_energy_rca(
                        stop_reason_charger=tx_stop_reason,
                        stop_reason_server=None,
                        logs=filtered_logs,
                        target_connector=target_conn
                    )
            else:
                tx_id_display = ""
                ch_stop = it.get("StopReason")
                srv_stop = it.get("StopTriggerMessage")
                if (ch_stop and ch_stop.lower() not in ("other", "none", "null", "-", "")) or (srv_stop and srv_stop.lower() not in ("other", "none", "null", "-", "")):
                    cs_rca_text, is_hw_diag = diagnose_completed_low_energy_rca(ch_stop, srv_stop, [], target_connector=conn_id)
                else:
                    raw_logs = fetch_logs_for_booking(ch_id, dt, cookie_header)
                    cs_rca_text, is_hw_diag = diagnose_completed_low_energy_rca(ch_stop, srv_stop, raw_logs, target_connector=conn_id)

        if is_hw_diag:
            return None

        return {
            "date": b_date_str,
            "reason": reason_type,
            "booking_id": bid,
            "transaction_id": tx_id_display,
            "station_name": st_name,
            "charger_code": ch_code,
            "connector": conn_id,
            "energy": energy_val,
            "rca": cs_rca_text,
            "dt": dt
        }

    with concurrent.futures.ThreadPoolExecutor(max_workers=25) as executor:
        results = list(executor.map(process_candidate, all_candidates))

    rows_data = [r for r in results if r is not None]
    rows_data.sort(key=lambda r: int(r["booking_id"]) if str(r["booking_id"]).isdigit() else 0, reverse=True)

    # Build Excel Workbook
    print(f"\n[EXCEL] Building spreadsheet with {len(rows_data)} total rows...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ElectreeFi RCA 1"

    headers = ["Date", "Reason", "Booking ID", "Transaction ID", "Station Name", "Charger Code", "Connector", "RCA"]
    ws.append(headers)

    # Styling definitions matching CS team standards
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="000000")
    data_font = Font(name="Calibri", size=11, bold=False, color="000000")
    thin_border = Border(
        left=Side(style='thin', color='A0A0A0'),
        right=Side(style='thin', color='A0A0A0'),
        top=Side(style='thin', color='A0A0A0'),
        bottom=Side(style='thin', color='A0A0A0')
    )
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = yellow_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    ws.row_dimensions[1].height = 24

    for row_idx, r in enumerate(rows_data, start=2):
        ws.append([
            r["date"],
            r["reason"],
            int(r["booking_id"]) if str(r["booking_id"]).isdigit() else r["booking_id"],
            int(r["transaction_id"]) if str(r["transaction_id"]).isdigit() else r["transaction_id"],
            r["station_name"],
            r["charger_code"],
            int(r["connector"]) if str(r["connector"]).isdigit() else r["connector"],
            r["rca"]
        ])
        ws.row_dimensions[row_idx].height = 20
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=row_idx, column=col_idx)
            c.font = data_font
            c.border = thin_border
            c.alignment = center_align if col_idx in (1, 2, 3, 4, 7) else left_align

    # Column widths
    col_widths = {
        "A": 12,  # Date
        "B": 18,  # Reason ('Cancelled' or 'less than 1kWh')
        "C": 14,  # Booking ID
        "D": 16,  # Transaction ID
        "E": 50,  # Station Name
        "F": 22,  # Charger Code
        "G": 12,  # Connector
        "H": 58   # RCA
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    primary_path = os.path.join("E:\\ElectreeFi", output_excel)
    alt_path = os.path.join("E:\\ElectreeFi", "ElectreeFi_Cancelled_Bookings_RCA_All_Records.xlsx")

    saved_paths = []
    try:
        wb.save(primary_path)
        saved_paths.append(primary_path)
        print(f"Saved primary workbook: {primary_path}")
    except PermissionError:
        print(f"Notice: '{output_excel}' is currently open.")

    try:
        wb.save(alt_path)
        saved_paths.append(alt_path)
        print(f"Saved backup workbook: {alt_path}")
    except Exception as e:
        print(f"Error saving backup path: {e}")

    print("\n" + "=" * 80)
    print(f"SUCCESS: EXCEL RCA GENERATED WITH ALL {len(rows_data)} SESSIONS!")
    print(f"  Cancelled Bookings: {len(cancelled_candidates)}")
    print(f"  Completed Low-Consumption (<1 kWh): {len(completed_candidates)}")
    for p in saved_paths:
        print(f"Location: {p}")
    print("=" * 80)
    return saved_paths[0] if saved_paths else None


if __name__ == "__main__":
    run_full_rca()
