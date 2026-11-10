"""Use the unchanged PDF review pipeline with the portable local layout."""

import sys
import tempfile
from pathlib import Path

import pdfs_review as core
from pdfs_settings import settings_args

SUPPORT = Path(__file__).resolve().parent
ROOT = SUPPORT.parent


def main():
    try:
        if sys.argv[1:]:
            raise ValueError(
                "Edit pdfs-review.txt; the portable entry point accepts no scan arguments"
            )
        args = settings_args(ROOT / "pdfs-review.txt")
        core.PROJECT_DIR = ROOT
        core.SUPPORT_DIR = SUPPORT
        backend = SUPPORT / "state/backend/7z.exe"
        if not backend.is_file():
            raise FileNotFoundError("Verified archive backend is missing; extract a fresh ZIP")
        args.extend(["--seven-zip", str(backend)])
        core.rarfile.SEVENZIP_TOOL = str(backend)
        core.rarfile.tool_setup(
            force=True, unrar=False, unar=False, bsdtar=False, sevenzip=True, sevenzip2=False
        )
        with tempfile.TemporaryDirectory(prefix="work-", dir=SUPPORT / "state/temp") as temporary:
            core.RUNTIME_DIR = Path(temporary)
            return core.run(args)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
