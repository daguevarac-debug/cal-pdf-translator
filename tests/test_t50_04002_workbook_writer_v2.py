from cal_translator.excel.com_backend import retry_com_call
from cal_translator.formats.t50_04002.workbook_writer_v2 import channel_result_matrix


class _RetryableComError(Exception):
    hresult = -2147418111


def _payload() -> dict:
    results = []
    for channel in "ABC":
        for index in range(24):
            results.append(
                {
                    "channel": channel,
                    "range": "1,0 Ω / 45 A",
                    "specified_value": f"{index},000 mΩ",
                    "average_measured_value": f"{index},001 mΩ",
                    "bias": "0,001 mΩ",
                    "maximum_permissible_error": None,
                    "expanded_uncertainty": "0,005 mΩ",
                    "coverage_factor": "2,0",
                    "cmc": None,
                    "pass_fail": None,
                    "tur": None,
                    "tar": None,
                }
            )
    return {"results": results}


def test_channel_matrix_is_24_by_11_and_suppresses_repeated_range() -> None:
    matrix = channel_result_matrix(_payload(), "A")

    assert len(matrix) == 24
    assert all(len(row) == 11 for row in matrix)
    assert matrix[0][0] == "1,0 Ω / 45 A"
    assert matrix[1][0] == ""
    assert matrix[0][1] == "0,000 mΩ"


def test_retry_com_call_recovers_from_transient_rejections() -> None:
    attempts = 0

    def operation() -> str:
        nonlocal attempts
        attempts += 1
        if attempts < 3:
            raise _RetryableComError("busy")
        return "ok"

    assert retry_com_call(operation, attempts=4, initial_delay=0) == "ok"
    assert attempts == 3
