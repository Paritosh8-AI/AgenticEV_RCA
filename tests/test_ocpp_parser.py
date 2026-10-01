"""
Unit tests for OCPP-J message parser.
"""

from src.rca.ocpp_parser import parse_ocpp_message


def test_parse_call_stop_transaction():
    raw = '[2, "uuid-123", "StopTransaction", {"transactionId": 999, "reason": "EVDisconnected", "connectorId": 2}]'
    entry = parse_ocpp_message(raw_message=raw)

    assert entry.message_id == "uuid-123"
    assert entry.parsed_action == "StopTransaction"
    assert entry.reason == "EVDisconnected"
    assert entry.connector_id == 2
    assert entry.error_code is None


def test_parse_call_status_notification():
    raw = '[2, "uuid-456", "StatusNotification", {"connectorId": 1, "errorCode": "GroundFailure", "status": "Faulted"}]'
    entry = parse_ocpp_message(raw_message=raw)

    assert entry.message_id == "uuid-456"
    assert entry.parsed_action == "StatusNotification"
    assert entry.error_code == "GroundFailure"
    assert entry.connector_id == 1


def test_parse_raw_json_dict_payload():
    raw = '{"connectorId": 3, "reason": "EmergencyStop"}'
    entry = parse_ocpp_message(raw_message=raw, event_name="StopTransaction")

    assert entry.parsed_action == "StopTransaction"
    assert entry.connector_id == 3
    assert entry.reason == "EmergencyStop"


def test_parse_malformed_string_fallback():
    raw = 'Charger stopped with connectorId: 2, reason=PowerLoss'
    entry = parse_ocpp_message(raw_message=raw)

    assert entry.connector_id == 2
    assert entry.reason == "PowerLoss"
