"""
Unit tests for Real-Time Live CMS Telemetry Client & TelemetryCopilot
"""

import pytest
from unittest.mock import MagicMock, patch
from src.rca.live_cms_client import LiveCMSClient
from src.rca.telemetry_copilot import TelemetryCopilot


def test_live_cms_client_cookie_header(tmp_path):
    session_file = tmp_path / "test_session.json"
    session_file.write_text('{"cookies": [{"name": "ASP.NET_SessionId", "value": "test12345", "domain": "emonitoring.electreefi.com"}]}', encoding="utf-8")
    client = LiveCMSClient(session_path=str(session_file))
    hdr = client.get_cookie_header()
    assert "ASP.NET_SessionId=test12345" in hdr


def test_live_cms_client_unauthenticated_fallback():
    client = LiveCMSClient(session_path="non_existent_file.json")
    assert client.get_cookie_header() == ""
    assert client.is_authenticated() is False
    res = client.fetch_live_charger_statuses()
    assert res["total_sampled"] == 0
    assert "error" in res


def test_copilot_live_query_detection():
    copilot = TelemetryCopilot()
    assert copilot._is_live_query("What is the live network pulse right now?") is True
    assert copilot._is_live_query("Who is charging right now?") is True
    assert copilot._is_live_query("Show me bookings cancelled today") is True
    assert copilot._is_live_query("Which stations have the most cable lock failures?") is False


def test_copilot_live_pulse_routing():
    copilot = TelemetryCopilot()
    mock_pulse = {
        "is_live": True,
        "is_authenticated": True,
        "timestamp": "2026-10-03 13:00:00",
        "chargers": {"total_sampled": 100, "status_counts": {"Active": 75, "Closed": 25}, "station_count": 40},
        "sessions": {"active_count": 12, "total_kwh_delivered": 150.5, "completed_today_count": 10},
        "cancellations": {"total_cancelled_today": 8, "direct_cancelled_count": 5, "roaming_cancelled_count": 3, "top_reasons": {"Invalid Session": 4}},
        "connectors": {"total_connectors": 50, "available_connectors": 30, "power_types": {"DC": 30, "AC": 20}}
    }

    with patch.object(copilot.live_client, "is_authenticated", return_value=True), \
         patch.object(copilot.live_client, "get_full_live_network_pulse", return_value=mock_pulse):
        
        res = copilot.ask("What is the live network pulse right now?")
        assert "Live CMS Real-Time Network Pulse" in res["answer"]
        assert any(m["label"] == "Active Chargers (Live)" for m in res["metrics"])
        assert res["chart_data"]["type"] == "bar"


def test_copilot_live_sessions_routing():
    copilot = TelemetryCopilot()
    mock_sessions = {
        "is_live": True,
        "timestamp": "2026-10-03 13:00:00",
        "active_count": 2,
        "completed_today_count": 5,
        "total_kwh_delivered": 45.2,
        "active_sessions": [
            {
                "session_id": "sess-12345",
                "party": "Yo charge",
                "source_party": "IOC",
                "connector_id": "A",
                "status": "ACTIVE",
                "kwh": 22.5,
                "current_soc": 78,
                "total_cost": "450.00"
            }
        ]
    }

    with patch.object(copilot.live_client, "is_authenticated", return_value=True), \
         patch.object(copilot.live_client, "fetch_live_active_sessions", return_value=mock_sessions):
        
        res = copilot.ask("Who is charging right now?")
        assert "Real-Time Active Charging Sessions" in res["answer"]
        assert "sess-12345" in res["answer"]
        assert any(m["label"] == "Ongoing Charging" for m in res["metrics"])


def test_copilot_live_cancellations_routing():
    copilot = TelemetryCopilot()
    mock_cancels = {
        "is_live": True,
        "timestamp": "2026-10-03 13:00:00",
        "total_cancelled_today": 14,
        "direct_cancelled_count": 9,
        "roaming_cancelled_count": 5,
        "top_reasons": {"Canceled by Invalid Session": 8, "User Abort": 6},
        "sample_cancellations": [
            {"id": "7630225", "source": "Direct CMS Booking", "station": "Highway Plaza", "reason": "Solenoid Timeout"}
        ]
    }

    with patch.object(copilot.live_client, "is_authenticated", return_value=True), \
         patch.object(copilot.live_client, "fetch_live_cancellations_today", return_value=mock_cancels):
        
        res = copilot.ask("Show me bookings cancelled today")
        assert "Today's Cancelled Bookings & Aborts" in res["answer"]
        assert any(m["label"] == "Total Network Aborts" for m in res["metrics"])


def test_copilot_live_ocpi_cancellations_routing():
    copilot = TelemetryCopilot()
    mock_ocpi = {
        "is_live": True,
        "timestamp": "2026-10-03 13:00:00",
        "total_ocpi_cancelled": 73,
        "parties": {"IOC": 24, "VIN": 24, "MPC": 8},
        "reasons": {"Cancelled by Scheduler (Timeout)": 45, "Cancelled by User": 18, "Cancelled by Invalid Session": 10},
        "items": [
            {"booking_id": "163396", "session_id": "-", "party": "IOC", "user_name": "Vansh", "action_group": "Cancelled by User", "station": "Rajendra Station"}
        ]
    }

    with patch.object(copilot.live_client, "is_authenticated", return_value=True), \
         patch.object(copilot.live_client, "fetch_live_ocpi_cancellations_today", return_value=mock_ocpi):
        
        res = copilot.ask("WHAT IS THE STATUS OF OCPI CANCELLED SESSIONS FOR TODAY?")
        assert "Live OCPI Roaming Cancelled Reservations (73 Today)" in res["answer"]
        assert "tab=CancelledReservation" in res["answer"]
        assert any(m["label"] == "OCPI Cancelled Today" and m["value"] == "73 items" for m in res["metrics"])
        assert any(m["label"] == "Grid Counter" and m["value"] == "1 - 73 of 73" for m in res["metrics"])
        assert "163396" in res["answer"]


def test_copilot_live_direct_cancellations_routing():
    copilot = TelemetryCopilot()
    mock_direct = {
        "is_live": True,
        "timestamp": "2026-10-03 13:00:00",
        "total_direct_cancelled": 157,
        "reasons": {"User / Remote Abort": 157},
        "stations": {"Highway Station": 12},
        "items": [
            {"booking_id": "7630225", "station": "Highway Station", "charger": "DEL01", "user": "Driver", "reason": "User / Remote Abort"}
        ]
    }

    with patch.object(copilot.live_client, "is_authenticated", return_value=True), \
         patch.object(copilot.live_client, "fetch_live_direct_cancellations_today", return_value=mock_direct):
        
        res = copilot.ask("Show Direct CMS cancellations")
        assert "Live Direct CMS Cancelled Bookings (157 Today)" in res["answer"]
        assert any(m["label"] == "Direct Cancelled Today" and m["value"] == "157 bookings" for m in res["metrics"])
        assert any(m["label"] == "Grid Counter" and m["value"] == "1 - 157 of 157" for m in res["metrics"])



def test_copilot_live_connectors_routing():
    copilot = TelemetryCopilot()
    mock_conn = {
        "is_live": True,
        "timestamp": "2026-10-03 13:00:00",
        "total_connectors": 40,
        "available_connectors": 24,
        "power_types": {"DC": 25, "AC": 15},
        "sample_guns": [
            {"station": "Miyapur Metro", "city": "Hyderabad", "type": "CCS2", "power": "DC", "price_per_kwh": "14.90", "status": "AVAILABLE"}
        ]
    }

    with patch.object(copilot.live_client, "is_authenticated", return_value=True), \
         patch.object(copilot.live_client, "fetch_live_connectors_summary", return_value=mock_conn):
        
        res = copilot.ask("What are the live connector tariffs and pricing?")
        assert "Live Connector Availability & Tariff Rates" in res["answer"]
        assert any(m["label"] == "Available Guns" for m in res["metrics"])


def test_copilot_live_escalation_email():
    copilot = TelemetryCopilot()
    mock_pulse = {
        "timestamp": "2026-10-03 13:00:00",
        "chargers": {"total_sampled": 100, "status_counts": {"Active": 70, "Closed": 30}},
        "sessions": {"active_count": 10},
        "cancellations": {"total_cancelled_today": 25, "top_reasons": {"Contactor Trip": 15}},
        "connectors": {"available_connectors": 20}
    }

    with patch.object(copilot.live_client, "is_authenticated", return_value=True), \
         patch.object(copilot.live_client, "get_full_live_network_pulse", return_value=mock_pulse):
        
        res = copilot.ask("Draft an engineering escalation email for today's issues")
        assert "Generated Live Escalation Notice" in res["answer"]
        assert "25 Aborted Sessions Detected Today" in res["answer"]
