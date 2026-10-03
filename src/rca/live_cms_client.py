"""
ElectreeFi CMS Live Telemetry Client
Directly connects to ElectreeFi CMS & Partner CPO Portals via authenticated HTTP/REST
to retrieve real-time EV charging telemetry, active sessions, charger health,
connectors, and cancellation events without reliance on static files.
"""

import os
import json
import time
import urllib.request
import urllib.parse
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional

from src.config import SESSION_STORAGE_PATH

class LiveCMSClient:
    def __init__(self, session_path: Optional[str] = None):
        self.session_path = session_path or str(SESSION_STORAGE_PATH)
        self.base_url = "https://emonitoring.electreefi.com"
        self._cache: Dict[str, Any] = {}
        self._cache_ttls: Dict[str, float] = {}
        self.default_ttl = 45.0  # 45 seconds cache to balance real-time freshness and API latency

    def get_cookie_header(self) -> str:
        """Retrieves and formats session cookies from the saved session state file."""
        if not os.path.exists(self.session_path):
            return ""
        try:
            with open(self.session_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            cookies = data.get("cookies", [])
            valid_cookies = [f"{c['name']}={c['value']}" for c in cookies if "electreefi.com" in c.get("domain", "")]
            return "; ".join(valid_cookies)
        except Exception:
            return ""

    def is_authenticated(self) -> bool:
        """Fast check if the current session cookie exists and is valid on the CMS."""
        cookie_header = self.get_cookie_header()
        if not cookie_header:
            return False

        cache_key = "auth_check"
        now = time.time()
        if cache_key in self._cache and (now - self._cache_ttls.get(cache_key, 0) < 60):
            return self._cache[cache_key]

        try:
            today = datetime.now().strftime("%Y-%m-%d")
            url = f"{self.base_url}/ChargingStationManagement/AdminBookingDetails/CancelledBookingDataThroughAjax?StationId=&StartDate={today}&Enddate={today}&stateId=&cityId=&transactionTypeId=-1&Id=0&vin=&chargerCode=&user=&vehicleNumber="
            headers = {
                "X-Requested-With": "XMLHttpRequest",
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "Cookie": cookie_header
            }
            payload = urllib.parse.urlencode({"sort": "", "page": "1", "pageSize": "1", "group": "", "filter": ""}).encode("utf-8")
            req = urllib.request.Request(url, data=payload, headers=headers)
            with urllib.request.urlopen(req, timeout=6) as res:
                text = res.read().decode("utf-8", errors="replace")
                is_valid = ("Data" in text and "/Account/Login" not in text)
                self._cache[cache_key] = is_valid
                self._cache_ttls[cache_key] = now
                return is_valid
        except Exception:
            self._cache[cache_key] = False
            self._cache_ttls[cache_key] = now
            return False

    def _post_ajax(self, endpoint: str, params: Optional[Dict[str, str]] = None, payload: Optional[Dict[str, str]] = None, timeout: float = 12.0) -> Dict[str, Any]:
        """Performs an authenticated POST request against an AJAX Kendo UI endpoint."""
        cookie_header = self.get_cookie_header()
        if not cookie_header:
            return {"Data": [], "Total": 0, "error": "Not authenticated. No session cookies found."}

        full_url = f"{self.base_url}{endpoint}" if endpoint.startswith("/") else f"{self.base_url}/{endpoint}"
        if params:
            full_url += ("?" if "?" not in full_url else "&") + urllib.parse.urlencode(params)

        default_payload = {"sort": "", "page": "1", "pageSize": "50", "group": "", "filter": ""}
        if payload:
            default_payload.update(payload)

        headers = {
            "X-Requested-With": "XMLHttpRequest",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Cookie": cookie_header,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
        }

        try:
            req_data = urllib.parse.urlencode(default_payload).encode("utf-8")
            req = urllib.request.Request(full_url, data=req_data, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as res:
                content = res.read().decode("utf-8", errors="replace")
                if "/Account/Login" in content:
                    return {"Data": [], "Total": 0, "error": "Session expired or redirected to login."}
                return json.loads(content)
        except Exception as e:
            return {"Data": [], "Total": 0, "error": str(e)}

    # -------------------------------------------------------------------------
    # High-Level Real-Time Telemetry Methods
    # -------------------------------------------------------------------------

    def fetch_live_charger_statuses(self, page_size: int = 100, force_refresh: bool = False) -> Dict[str, Any]:
        """
        Fetches live charger health, connectivity status, and heartbeat across the network.
        Returns active, closed, charging, faulted counts and station-level groupings.
        """
        cache_key = f"live_chargers_{page_size}"
        now = time.time()
        if not force_refresh and cache_key in self._cache and (now - self._cache_ttls.get(cache_key, 0) < self.default_ttl):
            return self._cache[cache_key]

        res = self._post_ajax(
            "/OCPPManagement/OCPP/LoadChargerStatusViewThroughAjaxFor_Data/3",
            payload={"page": "1", "pageSize": str(page_size), "filter": ""}
        )

        items = res.get("Data", [])
        status_counts = {"Active": 0, "Closed": 0, "Charging": 0, "Faulted": 0, "Unavailable": 0, "Suspended": 0, "Other": 0}
        stations_map = {}
        sample_chargers = []

        for it in items:
            raw_status = (it.get("Status") or "Unknown").strip()
            norm_status = raw_status.capitalize()
            if "Active" in norm_status or "Available" in norm_status or "Operat" in norm_status:
                status_counts["Active"] += 1
            elif "Charg" in norm_status:
                status_counts["Charging"] += 1
            elif "Close" in norm_status or "Offline" in norm_status:
                status_counts["Closed"] += 1
            elif "Fault" in norm_status or "Error" in norm_status or "Inoperat" in norm_status:
                status_counts["Faulted"] += 1
            elif "Suspend" in norm_status:
                status_counts["Suspended"] += 1
            elif "Unavail" in norm_status:
                status_counts["Unavailable"] += 1
            else:
                status_counts["Other"] += 1

            stn = it.get("ChargingStationName") or "Central Station"
            if stn not in stations_map:
                stations_map[stn] = {"total": 0, "active": 0, "closed": 0, "faulted": 0, "chargers": []}
            stations_map[stn]["total"] += 1
            if "Active" in norm_status:
                stations_map[stn]["active"] += 1
            elif "Close" in norm_status:
                stations_map[stn]["closed"] += 1
            elif "Fault" in norm_status:
                stations_map[stn]["faulted"] += 1

            code = it.get("ChargerCode") or it.get("ConnectionID") or "N/A"
            stations_map[stn]["chargers"].append(code)

            if len(sample_chargers) < 25:
                sample_chargers.append({
                    "code": code,
                    "station": stn,
                    "status": raw_status,
                    "type": it.get("ChargerType") or "DC Fast",
                    "heartbeat": it.get("LastHeartBeat") or "Live",
                    "charger_id": it.get("ChargerID")
                })

        summary = {
            "is_live": True,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_sampled": len(items),
            "status_counts": status_counts,
            "station_count": len(stations_map),
            "stations": stations_map,
            "sample_chargers": sample_chargers,
            "error": res.get("error")
        }

        self._cache[cache_key] = summary
        self._cache_ttls[cache_key] = now
        return summary

    def fetch_live_active_sessions(self, force_refresh: bool = False) -> Dict[str, Any]:
        """
        Fetches live ongoing / active charging sessions across OCPI Roaming network.
        Identifies who is currently charging, delivered kWh, current SOC %, cost, and duration.
        """
        cache_key = "live_active_sessions"
        now = time.time()
        if not force_refresh and cache_key in self._cache and (now - self._cache_ttls.get(cache_key, 0) < self.default_ttl):
            return self._cache[cache_key]

        today = datetime.now().strftime("%Y-%m-%d")
        res = self._post_ajax(
            "/Roaming/OCPISession/LoadOcpiSessionGridViewThroughAjax",
            params={"StartDate": today, "Enddate": today},
            payload={"page": "1", "pageSize": "50"}
        )

        items = res.get("Data", [])
        active_sessions = []
        completed_today = []
        total_kwh_delivered = 0.0

        for it in items:
            status = (it.get("Status") or "UNKNOWN").upper()
            kwh = 0.0
            try:
                kwh = float(it.get("Kwh") or 0.0)
            except Exception:
                pass
            total_kwh_delivered += kwh

            session_record = {
                "session_id": it.get("SessionId"),
                "party": it.get("CredentialPartyName") or it.get("PartyId") or "OCPI",
                "source_party": it.get("SourcePartyName") or it.get("SourcePartyId") or "CPO",
                "hub_partner": it.get("HubPartyName") or "Direct",
                "evse_uid": it.get("EvseUID"),
                "connector_id": it.get("ConnectorId"),
                "status": status,
                "start_time": it.get("StartDatetime") or it.get("CreatedOn"),
                "kwh": round(kwh, 2),
                "initial_soc": it.get("InitialSOC"),
                "current_soc": it.get("SOC"),
                "total_cost": it.get("TotalCost"),
                "currency": it.get("Currency") or "INR",
                "last_updated": it.get("LastUpdated") or it.get("ModifiedOn")
            }

            if status in ["ACTIVE", "PENDING", "IN_PROGRESS"]:
                active_sessions.append(session_record)
            else:
                completed_today.append(session_record)

        summary = {
            "is_live": True,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "active_count": len(active_sessions),
            "completed_today_count": len(completed_today),
            "total_kwh_delivered": round(total_kwh_delivered, 2),
            "active_sessions": active_sessions,
            "completed_sessions": completed_today[:15],
            "error": res.get("error")
        }

        self._cache[cache_key] = summary
        self._cache_ttls[cache_key] = now
        return summary

    def fetch_live_cancellations_today(self, force_refresh: bool = False) -> Dict[str, Any]:
        """
        Fetches today's aborted / cancelled bookings across regular CMS and OCPI Roaming.
        Extracts exact stop reasons, failure attribution, and impacted stations.
        """
        cache_key = "live_cancellations_today"
        now = time.time()
        if not force_refresh and cache_key in self._cache and (now - self._cache_ttls.get(cache_key, 0) < self.default_ttl):
            return self._cache[cache_key]

        today = datetime.now().strftime("%Y-%m-%d")

        # 1. Regular Bookings Cancelled
        reg_res = self._post_ajax(
            "/ChargingStationManagement/AdminBookingDetails/CancelledBookingDataThroughAjax",
            params={
                "StationId": "",
                "StartDate": today,
                "Enddate": today,
                "stateId": "",
                "cityId": "",
                "transactionTypeId": "-1",
                "Id": "0",
                "vin": "",
                "chargerCode": "",
                "user": "",
                "vehicleNumber": ""
            },
            payload={"page": "1", "pageSize": "50"}
        )
        reg_items = reg_res.get("Data", [])

        # 2. OCPI Roaming Reservations Cancelled
        ocpi_res = self._post_ajax(
            "/Roaming/OCPIReservation/LoadOcpiCancelledReservationGridViewThroughAjax",
            params={"StartDate": today, "Enddate": today, "transactionTypeId": "-1"},
            payload={"page": "1", "pageSize": "50"}
        )
        ocpi_items = ocpi_res.get("Data", [])

        cancellations = []
        failure_reasons = {}
        station_aborts = {}

        # Process Regular Cancelled
        for it in reg_items:
            b_id = str(it.get("ChargingStationBookingId") or "Unknown")
            stn = it.get("StationName") or "Central Station"
            reason = it.get("Reason") or it.get("Remark") or it.get("StopReason") or "User / Remote Abort"
            cancellations.append({
                "source": "Direct CMS Booking",
                "id": b_id,
                "station": stn,
                "charger": it.get("ChargerCode") or "N/A",
                "time": it.get("BookingInTime") or it.get("CreatedOn") or today,
                "reason": reason,
                "energy": it.get("EnergyConsumed_indecimal", 0.0)
            })
            failure_reasons[reason] = failure_reasons.get(reason, 0) + 1
            station_aborts[stn] = station_aborts.get(stn, 0) + 1

        # Process OCPI Cancelled
        for it in ocpi_items:
            party = it.get("PartyId") or "OCPI"
            sched_act = it.get("SchedularAction") or "Canceled by Invalid Session"
            stn = it.get("LocationName") or f"Roaming Station ({party})"
            cancellations.append({
                "source": f"Roaming ({party})",
                "id": str(it.get("ReservationId") or it.get("Id") or "ROAM-ABORT"),
                "station": stn,
                "charger": it.get("EvseUID") or "N/A",
                "time": it.get("CreatedOn") or today,
                "reason": sched_act,
                "energy": 0.0
            })
            failure_reasons[sched_act] = failure_reasons.get(sched_act, 0) + 1
            station_aborts[stn] = station_aborts.get(stn, 0) + 1

        summary = {
            "is_live": True,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_cancelled_today": len(cancellations),
            "direct_cancelled_count": len(reg_items),
            "roaming_cancelled_count": len(ocpi_items),
            "top_reasons": failure_reasons,
            "station_aborts": station_aborts,
            "sample_cancellations": cancellations[:25]
        }

        self._cache[cache_key] = summary
        self._cache_ttls[cache_key] = now
        return summary

    def fetch_live_connectors_summary(self, force_refresh: bool = False) -> Dict[str, Any]:
        """
        Fetches live connectors, charging guns, power types, format, tariffs, and status.
        """
        cache_key = "live_connectors"
        now = time.time()
        if not force_refresh and cache_key in self._cache and (now - self._cache_ttls.get(cache_key, 0) < self.default_ttl):
            return self._cache[cache_key]

        res = self._post_ajax(
            "/Roaming/OCPILocation/LoadLocationChargerConnectorsThroughAjax",
            params={"StatusId": "1"},
            payload={"page": "1", "pageSize": "60"}
        )

        items = res.get("Data", [])
        avail_count = 0
        occupied_count = 0
        power_types = {"DC": 0, "AC": 0}
        connector_types = {}
        sample_guns = []

        for it in items:
            st = (it.get("Status") or "AVAILABLE").upper()
            if "AVAIL" in st:
                avail_count += 1
            elif "OCCUPIED" in st or "CHARGING" in st:
                occupied_count += 1

            pt = (it.get("PowerType") or "DC").upper()
            power_types[pt] = power_types.get(pt, 0) + 1

            ct = it.get("ConnectorType") or "CCS2"
            connector_types[ct] = connector_types.get(ct, 0) + 1

            if len(sample_guns) < 15:
                sample_guns.append({
                    "station": it.get("LocationName"),
                    "city": it.get("City"),
                    "state": it.get("State"),
                    "type": ct,
                    "power": pt,
                    "format": it.get("Format") or "CABLE",
                    "price_per_kwh": it.get("Price") or "14.90",
                    "status": st,
                    "evse_id": it.get("EvseId"),
                    "party": it.get("PartyName") or it.get("PartyId")
                })

        summary = {
            "is_live": True,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_connectors": len(items),
            "available_connectors": avail_count,
            "occupied_connectors": occupied_count,
            "power_types": power_types,
            "connector_types": connector_types,
            "sample_guns": sample_guns,
            "error": res.get("error")
        }

        self._cache[cache_key] = summary
        self._cache_ttls[cache_key] = now
        return summary

    def get_full_live_network_pulse(self, force_refresh: bool = False) -> Dict[str, Any]:
        """
        Aggregates all live real-time streams into a unified high-level snapshot.
        """
        chargers = self.fetch_live_charger_statuses(page_size=100, force_refresh=force_refresh)
        sessions = self.fetch_live_active_sessions(force_refresh=force_refresh)
        cancellations = self.fetch_live_cancellations_today(force_refresh=force_refresh)
        connectors = self.fetch_live_connectors_summary(force_refresh=force_refresh)

        return {
            "is_live": True,
            "is_authenticated": self.is_authenticated(),
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "chargers": chargers,
            "sessions": sessions,
            "cancellations": cancellations,
            "connectors": connectors
        }


# Singleton instance
live_client = LiveCMSClient()
