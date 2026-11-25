# -*- coding: utf-8 -*-
"""从 wfp_version.py 生成 PyInstaller 的 --version-file。

版本号是单一来源：改 wfp_version.__version__ 一处即可，
数字版本取前三段，定制后缀（如 -eimitrilia.1）取末尾数字作为第 4 段。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from wfp_version import APP_NAME, __version__  # noqa: E402

COMPANY = "Eimitrilia Spotwinx"
DESCRIPTION = "Word 文档智能排版工具"
COPYRIGHT = "© 2026 Eimitrilia Spotwinx"


def numeric_parts(ver: str) -> tuple[int, int, int, int]:
    m = re.match(r"(\d+)\.(\d+)\.(\d+)(?:-[A-Za-z0-9]+\.(\d+))?", ver)
    if not m:
        return (0, 0, 0, 0)
    a, b, c = (int(x) for x in m.group(1, 2, 3))
    d = int(m.group(4) or 0)
    return (a, b, c, d)


def main() -> int:
    a, b, c, d = numeric_parts(__version__)
    out = ROOT / "version_info.txt"
    out.write_text(
        f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({a}, {b}, {c}, {d}),
    prodvers=({a}, {b}, {c}, {d}),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
        StringTable(
          '040904B0',
          [StringStruct('CompanyName', '{COMPANY}'),
           StringStruct('FileDescription', '{DESCRIPTION}'),
           StringStruct('FileVersion', '{__version__}'),
           StringStruct('InternalName', '{APP_NAME}'),
           StringStruct('LegalCopyright', '{COPYRIGHT}'),
           StringStruct('OriginalFilename', '{APP_NAME}.exe'),
           StringStruct('ProductName', '{APP_NAME}'),
           StringStruct('ProductVersion', '{__version__}')]
        )
      ]
    ),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
""",
        encoding="utf-8",
    )
    print(f"version_info.txt -> {out}  ({__version__} / {(a, b, c, d)})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
