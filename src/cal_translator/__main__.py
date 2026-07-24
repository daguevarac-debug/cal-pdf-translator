from __future__ import annotations

from cal_translator.formats.t50_04002.temperature_extractor_v2 import install_patch

install_patch()

from cal_translator.cli import main  # noqa: E402 - patch must load before CLI imports


if __name__ == "__main__":
    raise SystemExit(main())
