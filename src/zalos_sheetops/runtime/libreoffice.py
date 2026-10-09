import os
import platform
import shutil
import subprocess
from pathlib import Path


def recommended_install_command(system: str, distro_id: str | None) -> str:
    if system == "Darwin":
        return "brew install --cask libreoffice"
    if system != "Linux":
        return "Install LibreOffice and ensure soffice is on PATH"

    distro = (distro_id or "").lower()
    if distro in {"debian", "ubuntu"}:
        return (
            "sudo apt-get update && "
            "sudo apt-get install -y libreoffice libreoffice-calc"
        )
    if distro in {"fedora", "rhel", "centos", "amzn", "amazon"}:
        return "sudo dnf install -y libreoffice libreoffice-calc"
    if distro == "alpine":
        return "apk add --no-cache libreoffice"
    return "Install LibreOffice using your Linux distribution package manager"


def detect_libreoffice(
    executable: str | None = None,
    env: dict[str, str] | None = None,
) -> dict[str, object]:
    env_values = env if env is not None else os.environ
    candidates: list[tuple[str, str]] = []

    if executable:
        return _detect_candidate("explicit", executable, env_values)
    if env_values.get("SHEETOPS_SOFFICE"):
        return _detect_candidate("env", env_values["SHEETOPS_SOFFICE"], env_values)
    for name in ["soffice", "libreoffice"]:
        found = shutil.which(name)
        if found:
            candidates.append(("path", found))
    if platform.system() == "Darwin":
        candidates.append(
            ("common_path", "/Applications/LibreOffice.app/Contents/MacOS/soffice")
        )

    for source, candidate in candidates:
        result = _detect_candidate(source, candidate, env_values)
        if result["available"] or result.get("errorCode") == "libreoffice_broken":
            return result

    return _missing_result("", "auto")


def _detect_candidate(
    source: str,
    candidate: str,
    env_values: dict[str, str],
) -> dict[str, object]:
    path = Path(candidate).expanduser()
    if path.exists():
        probe = _probe_executable(path)
        if not probe["ok"]:
            return {
                "available": False,
                "errorCode": "libreoffice_broken",
                "executable": str(path),
                "source": source,
                "platform": platform.system().lower(),
                "stderr": probe["stderr"],
                "recommendedCommand": recommended_install_command(
                    platform.system(), _linux_id()
                ),
            }
        return {
            "available": True,
            "executable": str(path),
            "source": source,
            "platform": platform.system().lower(),
            "version": probe["version"],
        }
    return _missing_result(candidate, source)


def _missing_result(executable: str, source: str) -> dict[str, object]:
    return {
        "available": False,
        "errorCode": "libreoffice_missing",
        "executable": executable,
        "source": source,
        "platform": platform.system().lower(),
        "recommendedCommand": recommended_install_command(
            platform.system(), _linux_id()
        ),
    }


def _probe_executable(path: Path) -> dict[str, object]:
    try:
        completed = subprocess.run(
            [str(path), "--headless", "--version"],
            check=False,
            capture_output=True,
            text=True,
            timeout=10,
        )
    except (OSError, subprocess.TimeoutExpired) as error:
        return {"ok": False, "stderr": str(error), "version": ""}

    if completed.returncode != 0:
        return {
            "ok": False,
            "stderr": completed.stderr.strip() or completed.stdout.strip(),
            "version": "",
        }
    return {
        "ok": True,
        "stderr": "",
        "version": completed.stdout.strip() or completed.stderr.strip(),
    }


def _linux_id() -> str | None:
    path = Path("/etc/os-release")
    if not path.exists():
        return None
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        if line.startswith("ID="):
            return line.removeprefix("ID=").strip().strip('"')
    return None
