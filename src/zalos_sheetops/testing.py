"""In-process CLI runner for tests.

Replaces typer.testing.CliRunner so the CLI can stay dependency-free while
tests keep the same invoke(app, args) -> result API surface.
"""

import contextlib
import io


class Result:
    def __init__(self, exit_code: int, stdout: str, stderr: str) -> None:
        self.exit_code = exit_code
        self.stdout = stdout
        self.stderr = stderr
        self.output = stdout + stderr


class CliRunner:
    def invoke(
        self,
        app: object,
        args: list[str] | None = None,
        **_ignored: object,
    ) -> Result:
        stdout = io.StringIO()
        stderr = io.StringIO()
        exit_code = 0
        try:
            with (
                contextlib.redirect_stdout(stdout),
                contextlib.redirect_stderr(stderr),
            ):
                returned = app(list(args or []))
        except SystemExit as exc:
            code = exc.code
            if isinstance(code, int):
                exit_code = code
            elif code is not None:
                exit_code = 1
        else:
            exit_code = returned if isinstance(returned, int) else 0
        return Result(exit_code, stdout.getvalue(), stderr.getvalue())
