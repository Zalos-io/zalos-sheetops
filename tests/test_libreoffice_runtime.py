from pathlib import Path

from zalos_sheetops.runtime.libreoffice import (
    detect_libreoffice,
    recommended_install_command,
)


def test_recommended_install_command_for_macos() -> None:
    assert recommended_install_command("Darwin", None) == (
        "brew install --cask libreoffice"
    )


def test_recommended_install_command_for_ubuntu() -> None:
    assert recommended_install_command("Linux", "ubuntu") == (
        "sudo apt-get update && sudo apt-get install -y libreoffice libreoffice-calc"
    )


def test_recommended_install_command_for_fedora() -> None:
    assert recommended_install_command("Linux", "fedora") == (
        "sudo dnf install -y libreoffice libreoffice-calc"
    )


def test_recommended_install_command_for_alpine() -> None:
    assert recommended_install_command("Linux", "alpine") == (
        "apk add --no-cache libreoffice"
    )


def test_detect_libreoffice_reports_missing_explicit_executable() -> None:
    result = detect_libreoffice("/missing/soffice", env={})

    assert result["available"] is False
    assert result["errorCode"] == "libreoffice_missing"
    assert result["executable"] == "/missing/soffice"
    assert result["source"] == "explicit"


def test_detect_libreoffice_honors_env_executable(tmp_path: Path) -> None:
    executable = tmp_path / "soffice"
    executable.write_text("#!/bin/sh\nexit 0\n", encoding="utf-8")
    executable.chmod(0o755)

    result = detect_libreoffice(None, env={"SHEETOPS_SOFFICE": str(executable)})

    assert result["available"] is True
    assert result["executable"] == str(executable)
    assert result["source"] == "env"


def test_detect_libreoffice_reports_broken_executable(tmp_path: Path) -> None:
    executable = tmp_path / "soffice"
    executable.write_text(
        "#!/bin/sh\necho '/Applications/LibreOffice.app missing' >&2\nexit 1\n",
        encoding="utf-8",
    )
    executable.chmod(0o755)

    result = detect_libreoffice(str(executable), env={})

    assert result["available"] is False
    assert result["errorCode"] == "libreoffice_broken"
    assert result["executable"] == str(executable)
    assert result["source"] == "explicit"
    assert "LibreOffice.app missing" in result["stderr"]


def test_detect_libreoffice_probes_path_candidate(
    tmp_path: Path,
    monkeypatch,
) -> None:
    executable = tmp_path / "soffice"
    executable.write_text(
        "#!/bin/sh\necho 'broken wrapper' >&2\nexit 1\n",
        encoding="utf-8",
    )
    executable.chmod(0o755)
    monkeypatch.setattr(
        "zalos_sheetops.runtime.libreoffice.shutil.which",
        lambda name: str(executable) if name == "soffice" else None,
    )

    result = detect_libreoffice(None, env={})

    assert result["available"] is False
    assert result["errorCode"] == "libreoffice_broken"
    assert result["source"] == "path"
    assert "broken wrapper" in result["stderr"]
