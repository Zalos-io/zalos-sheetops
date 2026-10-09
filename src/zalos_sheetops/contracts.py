def success_response(
    command: str,
    data: dict[str, object],
    warnings: list[object] | None = None,
    artifacts: list[object] | None = None,
) -> dict[str, object]:
    return {
        "ok": True,
        "command": command,
        "data": data,
        "warnings": warnings or [],
        "artifacts": artifacts or [],
    }


def error_response(
    command: str,
    error_code: str,
    message: str,
    details: dict[str, object] | None = None,
    warnings: list[object] | None = None,
    artifacts: list[object] | None = None,
) -> dict[str, object]:
    return {
        "ok": False,
        "command": command,
        "errorCode": error_code,
        "message": message,
        "details": details or {},
        "warnings": warnings or [],
        "artifacts": artifacts or [],
    }
