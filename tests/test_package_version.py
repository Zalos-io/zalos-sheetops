from importlib.metadata import version

import zalos_sheetops


def test_dunder_version_matches_package_metadata() -> None:
    assert zalos_sheetops.__version__ == version("zalos-sheetops")
