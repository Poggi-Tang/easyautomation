"""Run the checked-out image CLI without installing or modifying Python packages."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from easy_uiauto.image_cli import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
