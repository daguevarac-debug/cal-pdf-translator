from cal_translator.formats.t50_04002.temperature_extractor_v2 import (
    extract_labeled_environmental_conditions,
)


def test_split_environmental_condition_rows_are_extracted() -> None:
    conditions = extract_labeled_environmental_conditions(
        """
        Temperatura máxima
        23,6 °C
        Humedad Relativa máxima
        44 %
        Temperatura mínima
        23,5 °C
        Humedad Relativa mínima
        44 %
        """
    )

    assert conditions is not None
    assert conditions.maximum_temperature == "23,6 °C"
    assert conditions.minimum_temperature == "23,5 °C"
    assert conditions.maximum_relative_humidity == "44 %"
    assert conditions.minimum_relative_humidity == "44 %"
