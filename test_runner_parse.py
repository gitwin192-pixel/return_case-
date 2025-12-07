import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent))

from runner import parse_refund_summary  # noqa: E402


def test_parse_refund_summary_basic():
    raw = {
        "error": 0,
        "data": {
            "exceptional_case_list": [
                {
                    "order_sn": "123",
                    "buyer": {"name": "buyer"},
                    "product_items": [
                        {
                            "product": {"name": "Widget", "sku": "SKU1"},
                            "model": {"name": "Red"},
                            "amount": 2,
                        }
                    ],
                    "request_solution_text": "Refund",
                    "request_reason_text": "Reason",
                    "header": {"status_text": "OK"},
                    "display_refund_amount": "100.00",
                    "forward_logistics_info": {
                        "shipping_carrier": "JNE",
                        "tracking_numbers": ["TRK1"],
                    },
                    "reverse_logistics_info": {
                        "shipping_carrier": "SPX",
                        "tracking_numbers": ["TRK2"],
                        "aggregated_logistics_status_text": "INTRANSIT",
                        "hint_text": "Hint",
                    },
                    "region": "ID",
                    "payment_method": "VA",
                }
            ]
        },
    }
    summary = parse_refund_summary(raw, "s2c", "store")
    assert summary["order_sn"] == "123"
    assert summary["buyer_name"] == "buyer"
    assert summary["product_name"] == "Widget (Red) [SKU1] x2"
    assert summary["product_sku"] == "SKU1"
    assert summary["qty"] == 2
    assert summary["forward_resi"] == "TRK1"
    assert summary["reverse_resi"] == "TRK2"
    assert summary["store_code"] == "s2c"
