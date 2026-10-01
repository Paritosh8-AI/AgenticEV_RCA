"""
Unit tests for RCA classification rules and correlation engine.
"""

from src.models import BookingDetails, OCPPLogEntry, RCAResult
from src.rca.mapper import classify_stop_reason, classify_error_code
from src.rca.engine import analyze_booking_rca, generate_rca_report
from src.rca.ocpp_parser import is_remote_start_status_rejected, is_remote_start_status_accepted
from generate_excel_rca_full import diagnose_cancelled_rca
from generate_excel_rca import map_to_cs_rca_text


def test_classify_stop_reasons():
    ev_disc = classify_stop_reason("EVDisconnected")
    assert ev_disc is not None
    assert ev_disc["category"] == "USER_ACTION"

    pwr_loss = classify_stop_reason("PowerLoss")
    assert pwr_loss is not None
    assert pwr_loss["category"] == "GRID_OR_POWER_FAULT"

    e_stop = classify_stop_reason("EmergencyStop")
    assert e_stop is not None
    assert e_stop["category"] == "SAFETY_TRIGGER"


def test_classify_error_codes():
    gf = classify_error_code("GroundFailure")
    assert gf is not None
    assert gf["category"] == "ELECTRICAL_SAFETY_FAULT"

    lock = classify_error_code("ConnectorLockFailure")
    assert lock is not None
    assert lock["category"] == "HARDWARE_FAULT"

    ev_comm = classify_error_code("EVCommunicationError")
    assert ev_comm is not None
    assert ev_comm["category"] == "EV_COMMUNICATION_FAULT"


def test_analyze_booking_rca_with_error_code():
    booking = BookingDetails(
        booking_id="BK-1001",
        charger_code="CH-01",
        connector_sequence_id=1,
        units_consumed_kw=0.2
    )
    logs = [
        OCPPLogEntry(
            message_id="m1",
            event_name="StatusNotification",
            event_type="",
            message_type="CALL",
            raw_message="",
            timestamp="2026-09-13 10:00",
            connector_id=1,
            error_code="ConnectorLockFailure"
        )
    ]

    result = analyze_booking_rca(booking, logs)
    assert result.booking_id == "BK-1001"
    assert result.category == "HARDWARE_FAULT"
    assert "Connector Solenoid Lock Failure" in result.root_cause
    assert result.confidence == "High"


def test_analyze_booking_rca_low_consumption_fallback():
    booking = BookingDetails(
        booking_id="BK-1002",
        charger_code="CH-02",
        connector_sequence_id=2,
        units_consumed_kw=0.4
    )
    result = analyze_booking_rca(booking, logs=[])
    assert result.category == "ABORTED_INITIATION"
    assert result.units_consumed_kw == 0.4


def test_generate_rca_report_aggregation():
    booking1 = BookingDetails(booking_id="BK-1", charger_code="CH-01", units_consumed_kw=0.2)
    booking2 = BookingDetails(booking_id="BK-2", charger_code="CH-01", units_consumed_kw=5.0)

    res1 = analyze_booking_rca(booking1, [])
    res2 = analyze_booking_rca(booking2, [])

    report = generate_rca_report("2026-09-13", [res1, res2])
    assert report.total_analyzed == 2
    assert report.low_consumption_count == 1
    assert report.cancelled_count == 1
    assert len(report.top_affected_chargers) == 1
    assert report.top_affected_chargers[0]["charger_code"] == "CH-01"


def test_remote_start_rejected_direct_payload():
    logs = [
        OCPPLogEntry(
            message_id="m1",
            event_name="RemoteStartTransaction",
            event_type="Response",
            message_type="CALLRESULT",
            raw_message='{"status": "Rejected"}',
            timestamp="2026-09-15 10:00:00",
            parsed_action="RemoteStartTransaction",
            parsed_payload={"status": "Rejected"},
            connector_id=1
        )
    ]
    assert is_remote_start_status_rejected(logs) is True
    rca_text, is_hw = diagnose_cancelled_rca(logs, target_connector=1)
    assert rca_text == "remote start transaction was rejected"
    assert is_hw is False


def test_remote_start_rejected_correlated_call_and_callresult():
    logs = [
        OCPPLogEntry(
            message_id="uuid-999",
            event_name="RemoteStartTransaction",
            event_type="Request",
            message_type="CALL",
            raw_message='[2, "uuid-999", "RemoteStartTransaction", {"connectorId": 1, "idTag": "TAG01"}]',
            timestamp="2026-09-15 10:00:00",
            parsed_action="RemoteStartTransaction",
            connector_id=1
        ),
        OCPPLogEntry(
            message_id="uuid-999",
            event_name="CALLRESULT",
            event_type="Response",
            message_type="CALLRESULT",
            raw_message='[3, "uuid-999", {"status": "Rejected"}]',
            timestamp="2026-09-15 10:00:01",
            parsed_payload={"status": "Rejected"},
            connector_id=1
        )
    ]
    assert is_remote_start_status_rejected(logs) is True
    rca_text, is_hw = diagnose_cancelled_rca(logs, target_connector=1)
    assert rca_text == "remote start transaction was rejected"


def test_remote_start_accepted_not_rejected():
    logs = [
        OCPPLogEntry(
            message_id="uuid-100",
            event_name="RemoteStartTransaction",
            event_type="Request",
            message_type="CALL",
            raw_message='[2, "uuid-100", "RemoteStartTransaction", {"connectorId": 1, "idTag": "TAG01"}]',
            timestamp="2026-09-15 10:00:00",
            parsed_action="RemoteStartTransaction",
            connector_id=1
        ),
        OCPPLogEntry(
            message_id="uuid-100",
            event_name="CALLRESULT",
            event_type="Response",
            message_type="CALLRESULT",
            raw_message='[3, "uuid-100", {"status": "Accepted"}]',
            timestamp="2026-09-15 10:00:01",
            parsed_payload={"status": "Accepted"},
            connector_id=1
        )
    ]
    assert is_remote_start_status_rejected(logs) is False
    assert is_remote_start_status_accepted(logs) is True
    rca_text, _ = diagnose_cancelled_rca(logs, target_connector=1)
    assert rca_text != "remote start transaction was rejected"


def test_high_temp_error_does_not_write_remote_start_rejected():
    logs = [
        OCPPLogEntry(
            message_id="m-temp",
            event_name="StatusNotification",
            event_type="Request",
            message_type="CALL",
            raw_message='[2, "m-temp", "StatusNotification", {"connectorId": 1, "errorCode": "HighTemperature", "status": "Faulted"}]',
            timestamp="2026-09-15 10:00:00",
            error_code="HighTemperature",
            connector_id=1
        )
    ]
    assert is_remote_start_status_rejected(logs) is False
    rca_text, _ = diagnose_cancelled_rca(logs, target_connector=1)
    assert "remote start transaction was rejected" not in rca_text
    assert "Remote start was rejected" not in rca_text
    assert rca_text == "Thermal Limit Exceeded"


def test_boot_notification_rejected_is_not_remote_start():
    logs = [
        OCPPLogEntry(
            message_id="boot-1",
            event_name="BootNotification",
            event_type="Response",
            message_type="CALLRESULT",
            raw_message='[3, "boot-1", {"status": "Rejected"}]',
            timestamp="2026-09-15 10:00:00",
            parsed_action="BootNotification",
            parsed_payload={"status": "Rejected"},
            connector_id=0
        )
    ]
    assert is_remote_start_status_rejected(logs) is False
    rca_text, _ = diagnose_cancelled_rca(logs, target_connector=1)
    assert rca_text != "remote start transaction was rejected"


def test_map_to_cs_rca_text_with_rejected_remote_start():
    logs = [
        OCPPLogEntry(
            message_id="m-rs",
            event_name="RemoteStartTransaction",
            event_type="Response",
            message_type="CALLRESULT",
            raw_message='{"status": "Rejected"}',
            timestamp="2026-09-15 10:00:00",
            parsed_payload={"status": "Rejected"}
        )
    ]
    rca_result = RCAResult(
        booking_id="BK-1",
        category="UNCLASSIFIED_VENDOR",
        root_cause="Vendor-Specific Other Reason",
        explanation="Charger reported Other",
        recommendation="Check logs"
    )
    result_text = map_to_cs_rca_text(rca_result, logs)
    assert result_text == "remote start transaction was rejected"


def test_map_to_cs_rca_text_without_rejected_remote_start():
    logs = []
    rca_result = RCAResult(
        booking_id="BK-2",
        category="UNCLASSIFIED_VENDOR",
        root_cause="Vendor-Specific Other Reason",
        explanation="Charger reported Other",
        recommendation="Check logs"
    )
    result_text = map_to_cs_rca_text(rca_result, logs)
    assert result_text != "remote start transaction was rejected"
    assert result_text != "Remote start was rejected"

