from __future__ import annotations

from cal_translator.formats.t50_04002.batch_diagnostics import (
    install_patch as install_batch_diagnostics_patch,
)
from cal_translator.formats.t50_04002.temperature_english_export_v2 import (
    install_patch as install_temperature_export_patch,
)
from cal_translator.formats.t50_04002.temperature_extractor_v2 import (
    install_patch as install_temperature_extractor_patch,
)
from cal_translator.formats.t50_04002.temperature_workbook_writer_v2 import (
    install_patch as install_temperature_writer_patch,
)

install_temperature_extractor_patch()
install_temperature_writer_patch()
install_temperature_export_patch()
install_batch_diagnostics_patch()

from cal_translator.cli import main  # noqa: E402 - patches must load before CLI imports


if __name__ == "__main__":
    raise SystemExit(main())
