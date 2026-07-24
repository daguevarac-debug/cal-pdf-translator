from cal_translator.models import CalibrationCertificate


def test_certificate_serializes_to_plain_dict() -> None:
    certificate = CalibrationCertificate(format_id="T50-04002", certificate_number="CAL-00000")
    payload = certificate.to_dict()
    assert payload["format_id"] == "T50-04002"
    assert payload["certificate_number"] == "CAL-00000"
    assert payload["results"] == []
