import hashlib
import json
from pathlib import Path


def test_fixture_manifest_entries_have_required_fields() -> None:
    path = Path("tests/fixtures/manifest.json")

    entries = json.loads(path.read_text())

    assert entries
    for entry in entries:
        assert set(entry) == {
            "fixtureId",
            "sourceRepo",
            "commit",
            "sourcePath",
            "license",
            "destination",
            "sha256",
            "behavior",
        }
        assert entry["fixtureId"]
        assert len(entry["commit"]) >= 7
        assert entry["license"]


def test_registered_upstream_fixtures_exist_and_match_sha256() -> None:
    entries = json.loads(Path("tests/fixtures/manifest.json").read_text())

    for entry in entries:
        path = Path(entry["destination"])
        assert path.exists(), entry["destination"]
        actual_hash = hashlib.sha256(path.read_bytes()).hexdigest()
        assert actual_hash == entry["sha256"], entry["fixtureId"]
