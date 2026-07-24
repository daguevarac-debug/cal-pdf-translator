from __future__ import annotations

# Install the corrected temperature implementations before importing any module
# that imports batch_processor. Otherwise batch_processor keeps stale function
# references captured before the patches are active.
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

# This module imports batch_processor, so it must load only after the corrected
# temperature functions have been installed.
from cal_translator.formats.t50_04002.batch_diagnostics import (  # noqa: E402
    install_patch as install_batch_diagnostics_patch,
)

install_batch_diagnostics_patch()

from cal_translator.cli import main  # noqa: E402 - controlled import order


if __name__ == "__main__":
    raise SystemExit(main())
