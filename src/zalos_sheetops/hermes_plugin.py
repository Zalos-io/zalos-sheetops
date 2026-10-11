"""Optional Hermes plugin: spreadsheet policy runs outside the model tool loop."""

import hashlib
import io
import json
import os
import re
import threading
from pathlib import Path

from openpyxl import Workbook

from zalos_sheetops.delivery import DeliveryPolicy

MAX_BYTES = 64 * 1024 * 1024


class SpreadsheetDelivery:
    def __init__(
        self,
        *,
        cwd: Path,
        engine: str | None,
        reader: object = None,
        writer: object = None,
        contract_name: str = ".sheetops-delivery.json",
    ) -> None:
        self.cwd = cwd.resolve()
        self.engine = engine
        self.reader = reader
        self.writer = writer
        self.contract_path = self.cwd / contract_name
        self.turns: dict[str, dict] = {}
        self._lock = threading.RLock()

    def _read(self, path: Path, limit: int = MAX_BYTES) -> bytes | None:
        if self.reader is not None:
            return self.reader(path, limit)
        try:
            with path.open("rb") as source:
                data = source.read(limit + 1)
            if len(data) > limit:
                raise ValueError("Workbook exceeds delivery limit")
            return data
        except FileNotFoundError:
            return None

    def prepare(self, turn_id: str, message: str) -> dict:
        with self._lock:
            paths = re.findall(
                r'"([^"\n]+\.xlsx)"|(?:@file:)?([^\s"\'<>]+\.xlsx)',
                message,
                flags=re.IGNORECASE,
            )
            sources = {}
            for quoted, token in paths:
                path = Path(quoted or token)
                path = (
                    (self.cwd / path).resolve()
                    if not path.is_absolute()
                    else path.resolve()
                )
                data = self._read(path)
                if data is not None:
                    sources[str(path)] = data
            state = {
                "sources": sources,
                "policy": None,
                "contract": None,
                "output": None,
                "error": "",
            }
            self.turns[turn_id] = state
            state["requires_output"] = bool(
                re.search(
                    r"\b(create|make|build|save|edit|update|fill|format|change|insert|delete)\b",
                    message,
                    re.I,
                )
                and (
                    paths
                    or re.search(r"\b(spreadsheet|workbook|excel)\b", message, re.I)
                )
            )
            if (sources or state["requires_output"]) and self._read(
                self.contract_path, 256 * 1024
            ) is not None:
                state["error"] = "A delivery contract existed before this turn."
            return {
                "active": True,
                "protectedRoots": [str(self.cwd)],
                "extensions": [".xlsx"],
                "inputArtifacts": [
                    {"path": path, "sha256": hashlib.sha256(data).hexdigest()}
                    for path, data in sources.items()
                ],
                "message": (
                    "Spreadsheet delivery is enforced by the runtime. Before editing, "
                    f"write {self.contract_path.name} containing "
                    "source (the input path), "
                    "output (a new .xlsx path), and allowedRanges (sheet names mapped "
                    "to bounded A1 ranges). For creation without an input, "
                    "omit source; "
                    "the runtime supplies an empty workbook with a Sheet worksheet. "
                    "Create the contract in a separate tool call "
                    "before creating output. "
                    "Scope then freezes. Delivery checks unrelated cells and evaluates "
                    "all formulas. Missing calculation engines block formula delivery."
                ),
            }

    def before_tool(self, turn_id: str) -> dict:
        self.observe(turn_id)
        state = self.turns.get(turn_id)
        if state is None or state["error"]:
            return {
                "action": "block",
                "message": state["error"]
                if state
                else "Delivery preparation is unavailable.",
            }
        return {"action": "allow"}

    def observe(self, turn_id: str) -> None:
        with self._lock:
            state = self.turns.get(turn_id)
            if state is None or state["error"]:
                return
            try:
                data = self._read(self.contract_path, 256 * 1024)
                if data is None:
                    return
                contract = json.loads(data)
                if state["contract"] is not None:
                    if contract != state["contract"]:
                        raise ValueError("The declared edit contract is frozen")
                    return
                source = contract.get("source")
                if source is None and len(state["sources"]) == 1:
                    source = next(iter(state["sources"]))
                if source is None and not state["sources"]:
                    stream = io.BytesIO()
                    Workbook().save(stream)
                    original = stream.getvalue()
                else:
                    source_path = (self.cwd / str(source)).resolve()
                    if str(source_path) not in state["sources"]:
                        raise ValueError(
                            "Source must be an original user-supplied workbook"
                        )
                    original = state["sources"][str(source_path)]
                    if self._read(source_path) != original:
                        raise ValueError(
                            "Source was edited before the contract was frozen"
                        )
                output = (self.cwd / contract["output"]).resolve()
                if (
                    not output.is_relative_to(self.cwd)
                    or output.suffix.lower() != ".xlsx"
                    or str(output) in state["sources"]
                ):
                    raise ValueError(
                        "Output must be a new .xlsx file inside the workspace"
                    )
                if self._read(output) is not None:
                    raise ValueError("Output existed before the contract was frozen")
                policy = DeliveryPolicy(original, engine=self.engine)
                try:
                    policy.declare(contract["allowedRanges"])
                except Exception:
                    policy.close()
                    raise
                state.update(policy=policy, contract=contract, output=output)
            except Exception as error:
                state["error"] = str(error)
                if state["policy"] is not None:
                    state["policy"]._receipt = None

    def validate(self, turn_id: str, response: str = "") -> dict:
        with self._lock:
            state = self.turns.get(turn_id)
            if (
                state is not None
                and state["contract"] is None
                and not state["requires_output"]
                and not state["sources"]
                and ".xlsx" not in response.lower()
            ):
                return {"action": "allow", "artifacts": []}
            self.observe(turn_id)
            state = self.turns.get(turn_id)
            if state is None:
                return {
                    "action": "block",
                    "message": "Delivery preparation is unavailable.",
                }
            claimed = re.findall(r'([^\s"\'<>()[\]]+\.xlsx)', response, flags=re.I)
            if (
                not state["error"]
                and state["policy"] is None
                and not state["requires_output"]
            ):
                unchanged = all(
                    self._read(Path(path)) == original
                    for path, original in state["sources"].items()
                )
                known = all(
                    str((self.cwd / path).resolve()) in state["sources"]
                    for path in claimed
                )
                if unchanged and known:
                    return {"action": "allow", "artifacts": []}
            if state["error"] or state["policy"] is None:
                return {
                    "action": "block",
                    "message": state["error"]
                    or "No edit contract was frozen before editing.",
                }
            try:
                output = self._read(state["output"])
                if output is None:
                    return {
                        "action": "block",
                        "message": "The declared output file is missing.",
                    }
                result = state["policy"].validate(output)
                if result["action"] == "allow":
                    delivered = state["policy"].validated_bytes
                    if delivered != output:
                        if self.writer is not None:
                            self.writer(state["output"], delivered)
                        else:
                            import tempfile

                            descriptor, temporary = tempfile.mkstemp(
                                dir=state["output"].parent
                            )
                            try:
                                with os.fdopen(descriptor, "wb") as handle:
                                    handle.write(delivered)
                                os.replace(temporary, state["output"])
                            finally:
                                Path(temporary).unlink(missing_ok=True)
                        if self._read(state["output"]) != delivered:
                            raise ValueError(
                                "Output changed while publishing evaluated bytes"
                            )
                if result["action"] == "block":
                    result["message"] += " " + json.dumps(
                        result.get("details", {}), separators=(",", ":")
                    )
                return result
            except Exception:
                state["policy"]._receipt = None
                return {
                    "action": "block",
                    "message": "Could not read the output for validation.",
                }

    def authorize(self, path: Path, data: bytes) -> dict:
        with self._lock:
            path = path.resolve()
            if path.suffix.lower() != ".xlsx" or not path.is_relative_to(self.cwd):
                return {"action": "allow"}
            for state in self.turns.values():
                if (
                    str(path) in state["sources"]
                    and hashlib.sha256(data).digest()
                    == hashlib.sha256(state["sources"][str(path)]).digest()
                ):
                    return {"action": "allow"}
                if (
                    path == state["output"]
                    and not state["error"]
                    and state["policy"].is_approved(data)
                ):
                    return {"action": "allow"}
            return {
                "action": "block",
                "message": "This spreadsheet has no approval for its current bytes.",
            }

    def close(self) -> None:
        for state in self.turns.values():
            if state["policy"] is not None:
                state["policy"].close()


def register(ctx: object) -> None:
    runtimes: dict[str, SpreadsheetDelivery] = {}
    pending: dict[tuple[str, str], SpreadsheetDelivery] = {}

    def prepare(
        task_id: str,
        turn_id: str,
        user_message: object,
        session_id: str = "",
        **kwargs: object,
    ) -> dict:
        from agent.message_content import flatten_message_text
        from tools.terminal_tool import _get_env_config
        from tools.terminal_tool_lifecycle import ensure_task_env

        message = flatten_message_text(user_message)
        config = _get_env_config()
        backend = None
        cwd = Path(config["cwd"])
        engine = ctx.get_config("calculation_engine")
        if engine is None:
            from hermes_platform.resolver.core import locate_command

            resolution = locate_command(
                "soffice",
                known_dirs=(
                    "/Applications/LibreOffice.app/Contents/MacOS",
                    "${PROGRAMFILES}/LibreOffice/program",
                    "/usr/lib/libreoffice/program",
                ),
            )
            engine = resolution.command[0] if resolution.found else None

        def read(path: Path, limit: int) -> bytes | None:
            nonlocal backend
            if backend is None and config["env_type"] != "local":
                backend = ensure_task_env(task_id)
            if backend is None:
                if config["env_type"] != "local":
                    raise RuntimeError("Terminal backend is unavailable")
                try:
                    with path.open("rb") as source:
                        data = source.read(limit + 1)
                    if len(data) > limit:
                        raise ValueError("File exceeds delivery limit")
                    return data
                except FileNotFoundError:
                    return None
            import tempfile

            from tools.environments.base import FileFetchError

            with tempfile.TemporaryDirectory(prefix="sheetops-fetch-") as temporary:
                destination = Path(temporary) / "file"
                try:
                    backend.fetch_file(str(path), str(destination), max_bytes=limit)
                except FileFetchError:
                    return None
                return destination.read_bytes()

        key = (session_id, str(cwd))
        previous = pending.get(key)
        resume = (
            previous is not None
            and ".xlsx" not in message.lower()
            and re.match(
                r"\s*(continue|fix|repair|retry|try again|finish)\b", message, re.I
            )
        )
        if resume:
            runtime = previous
            runtime.reader = read
            runtime.turns[turn_id] = next(reversed(runtime.turns.values()))
            result = {
                "active": True,
                "protectedRoots": [str(cwd)],
                "extensions": [".xlsx"],
                "inputArtifacts": [],
                "message": "Delivery retains the prior frozen contract.",
            }
        else:
            if previous is not None:
                pending.pop(key, None)
                previous.close()
            contract_name = (
                ".sheetops-delivery-"
                + hashlib.sha256(turn_id.encode()).hexdigest()[:20]
                + ".json"
            )
            runtime = SpreadsheetDelivery(
                cwd=cwd,
                engine=engine,
                reader=read,
                contract_name=contract_name,
            )
            result = runtime.prepare(turn_id, message)
        runtimes[turn_id] = runtime

        def host_path(value: str) -> str:
            if config["env_type"] == "docker":
                for volume in config["docker_volumes"]:
                    host, container, *_ = volume.split(":")
                    path = Path(value)
                    if path.is_relative_to(container):
                        return str(Path(host) / path.relative_to(container))
                if config["docker_mount_cwd_to_workspace"] and config["host_cwd"]:
                    return str(
                        Path(config["host_cwd"]) / Path(value).relative_to("/workspace")
                    )
                raise ValueError("Delivery requires a host-visible workspace mount")
            return value

        runtime.host_path = host_path

        def write(path: Path, data: bytes) -> None:
            import tempfile

            if config["env_type"] not in {"local", "docker"}:
                import base64
                import shlex

                script = """import base64,json,os,sys,tempfile
payload=json.load(sys.stdin)
target=os.path.realpath(payload['path'])
root=os.path.realpath(payload['root'])
if os.path.commonpath([root,target])!=root:
 raise ValueError('Output is outside the frozen workspace')
fd,temporary=tempfile.mkstemp(dir=os.path.dirname(target))
try:
 with os.fdopen(fd,'wb') as handle:
  handle.write(base64.b64decode(payload['data'],validate=True))
 os.replace(temporary,target)
finally:
 if os.path.exists(temporary): os.unlink(temporary)
"""
                result = backend.execute(
                    "python3 -c " + shlex.quote(script),
                    cwd=str(cwd),
                    timeout=10,
                    stdin_data=json.dumps(
                        {
                            "path": str(path),
                            "root": str(cwd),
                            "data": base64.b64encode(data).decode("ascii"),
                        }
                    ),
                )
                if result["returncode"] != 0:
                    raise ValueError("Could not publish evaluated workbook bytes")
                return
            target = Path(host_path(str(path))).resolve()
            if not target.is_relative_to(Path(host_path(str(cwd))).resolve()):
                raise ValueError("Output escaped the frozen workspace")
            descriptor, temporary = tempfile.mkstemp(dir=target.parent)
            try:
                with os.fdopen(descriptor, "wb") as handle:
                    handle.write(data)
                os.replace(temporary, target)
            finally:
                Path(temporary).unlink(missing_ok=True)

        runtime.writer = write
        runtime.pending_key = key
        result["protectedRoots"] = [
            host_path(path) for path in result["protectedRoots"]
        ]
        result["inputArtifacts"] = [
            {**artifact, "path": host_path(artifact["path"])}
            for artifact in result["inputArtifacts"]
        ]
        return result

    def before_tool(turn_id: str, tool_name: str = "", **kwargs: object) -> dict:
        runtime = runtimes.get(turn_id)
        if runtime is None:
            return {
                "action": "block",
                "message": "Delivery preparation is unavailable.",
            }
        if tool_name not in {
            "terminal",
            "write_file",
            "patch",
            "read_file",
            "delegate_task",
            "execute_code",
        }:
            return {"action": "allow"}
        return runtime.before_tool(turn_id)

    def validate(turn_id: str, response: object = "", **kwargs: object) -> dict:
        runtime = runtimes.get(turn_id)
        if runtime is None:
            return {
                "action": "block",
                "message": "Delivery preparation is unavailable.",
            }
        result = runtime.validate(turn_id, str(response or ""))
        if result["action"] == "allow":
            pending.pop(runtime.pending_key, None)
        else:
            pending[runtime.pending_key] = runtime
        if result["action"] == "allow" and runtime.turns[turn_id]["output"] is not None:
            result["artifacts"] = [
                {
                    "path": runtime.host_path(str(runtime.turns[turn_id]["output"])),
                    "sha256": result["sha256"],
                }
            ]
        return result

    def finish(turn_id: str, **kwargs: object) -> None:
        runtime = runtimes.pop(turn_id, None)
        if runtime is not None and runtime not in pending.values():
            runtime.close()

    ctx.register_hook("post_llm_call", finish)
    ctx.register_hook("prepare_delivery", prepare)
    ctx.register_hook("pre_tool_call", before_tool)
    ctx.register_hook("validate_delivery", validate)
