from __future__ import annotations

import hashlib
import tempfile
import unittest
import zipfile
from pathlib import Path

from scripts.build_portable import ROOT, build, payload


class BuildTests(unittest.TestCase):
    def test_payload_is_source_only_and_does_not_include_private_state(self):
        data = payload()
        self.assertIn("pdfs-review.bat", data)
        self.assertIn("support/pdfs_review.py", data)
        self.assertNotIn("folder_review.py", data)
        self.assertTrue(all(not name.endswith((".exe", ".dll", ".pyc")) for name in data))
        self.assertTrue(all("state/" not in name and ".local/" not in name for name in data))
        self.assertEqual(data["support/pdfs_review.py"], (ROOT / "pdfs_review.py").read_bytes())
        self.assertEqual(
            data["support/UPSTREAM_LICENSE.txt"],
            (ROOT / "portable/UPSTREAM_LICENSE.txt").read_bytes(),
        )

    def test_reproducible_archive_and_manifest(self):
        import json

        with tempfile.TemporaryDirectory() as temporary:
            first = build(Path(temporary) / "first")
            second = build(Path(temporary) / "second")
            self.assertEqual(
                hashlib.sha256(first.read_bytes()).digest(),
                hashlib.sha256(second.read_bytes()).digest(),
            )
            with zipfile.ZipFile(first) as archive:
                manifest = json.loads(archive.read("support/manifest.json"))
                self.assertNotIn("pdfs-review.txt", manifest["immutable_sha256"])
                self.assertEqual(len(archive.namelist()), len(manifest["immutable_sha256"]) + 2)
                for name, digest in manifest["immutable_sha256"].items():
                    self.assertEqual(hashlib.sha256(archive.read(name)).hexdigest(), digest)


if __name__ == "__main__":
    unittest.main()
