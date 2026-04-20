from __future__ import annotations

import logging
import os
import json
import shutil
import subprocess
import tempfile
from pathlib import Path


logger = logging.getLogger(__name__)


def _candidate_dotnet_commands() -> list[str]:
    candidates = []

    configured_dotnet = os.getenv("DOCUMENT_TOOL_DOTNET") or os.getenv("MATRIX_OPENXML_DOTNET")
    if configured_dotnet:
        candidates.append(configured_dotnet)

    path_dotnet = shutil.which("dotnet")
    if path_dotnet:
        candidates.append(path_dotnet)

    candidates.extend(
        [
            r"C:\Program Files\dotnet\dotnet.exe",
            r"C:\Program Files (x86)\dotnet\dotnet.exe",
            "/mnt/c/Program Files/dotnet/dotnet.exe",
            "/mnt/c/Program Files (x86)/dotnet/dotnet.exe",
        ]
    )

    return candidates


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _default_tool_path() -> str:
    return str(
        _repo_root()
        / "dotnet"
        / "MatrixOpenXmlTextboxTool"
        / "bin"
        / "Release"
        / "net8.0"
        / "MatrixOpenXmlTextboxTool.dll"
    )


def _default_project_path() -> str:
    return str(
        _repo_root()
        / "dotnet"
        / "MatrixOpenXmlTextboxTool"
        / "MatrixOpenXmlTextboxTool.csproj"
    )


def _default_tool_needs_rebuild(tool_path: str) -> bool:
    if os.path.abspath(tool_path) != os.path.abspath(_default_tool_path()):
        return False

    tool = Path(tool_path)
    if not tool.exists():
        return True

    source_paths = [
        Path(_default_project_path()),
        _repo_root() / "dotnet" / "MatrixOpenXmlTextboxTool" / "Program.cs",
    ]
    tool_mtime = tool.stat().st_mtime
    return any(path.exists() and path.stat().st_mtime > tool_mtime for path in source_paths)


def _find_dotnet_host() -> str | None:
    for dotnet_path in _candidate_dotnet_commands():
        if os.path.exists(dotnet_path) or shutil.which(dotnet_path):
            return dotnet_path
    return None


def _is_windows_dotnet_host(dotnet_path: str) -> bool:
    lowered = dotnet_path.lower().replace("\\", "/")
    return lowered.endswith("/dotnet.exe") or lowered.endswith("dotnet.exe")


def _running_under_wsl() -> bool:
    try:
        release = Path("/proc/sys/kernel/osrelease").read_text(errors="ignore").lower()
    except OSError:
        return False
    return "microsoft" in release or "wsl" in release


def _to_dotnet_path(path: str, dotnet_path: str) -> str:
    if not path or not _is_windows_dotnet_host(dotnet_path) or not _running_under_wsl():
        return path

    if not path.startswith("/"):
        return path

    process = subprocess.run(
        ["wslpath", "-w", path],
        capture_output=True,
        text=True,
        check=False,
    )
    converted_path = process.stdout.strip()
    return converted_path or path


def _build_default_tool_if_possible(tool_path: str) -> bool:
    if os.path.abspath(tool_path) != os.path.abspath(_default_tool_path()):
        return False

    project_path = _default_project_path()
    if not os.path.exists(project_path):
        logger.warning("Document tool project was not found: %s", project_path)
        return False

    dotnet_path = _find_dotnet_host()
    if not dotnet_path:
        logger.warning("Cannot build document tool because no dotnet host was found.")
        return False

    logger.info("Document tool DLL was not found; building %s", project_path)
    project_path_arg = os.path.relpath(project_path, _repo_root())
    process = subprocess.run(
        [
            dotnet_path,
            "build",
            project_path_arg,
            "-c",
            "Release",
        ],
        capture_output=True,
        text=True,
        check=False,
        env=_subprocess_env_with_wsl_bridge(),
        cwd=str(_repo_root()),
    )

    if process.stdout.strip():
        logger.info("Document tool build stdout:\n%s", process.stdout.strip())
    if process.stderr.strip():
        logger.warning("Document tool build stderr:\n%s", process.stderr.strip())

    if process.returncode != 0:
        logger.warning("Document tool build exited with code %s", process.returncode)
        return False

    return os.path.exists(tool_path)


def _resolve_tool_command() -> list[str] | None:
    tool_path = (
        os.getenv("DOCUMENT_TOOL_PATH")
        or os.getenv("SYNCFUSION_DOCUMENT_TOOL")
        or os.getenv("MATRIX_OPENXML_TEXTBOX_TOOL")
        or _default_tool_path()
    )

    if tool_path.endswith(".dll"):
        if _default_tool_needs_rebuild(tool_path):
            if not _build_default_tool_if_possible(tool_path):
                logger.warning("Document tool DLL was not found or could not be rebuilt: %s", tool_path)
                return None

        dotnet_path = _find_dotnet_host()
        if dotnet_path:
            return [dotnet_path, _to_dotnet_path(tool_path, dotnet_path)]

        logger.warning(
            "Document tool points to a .dll, but no dotnet host was found. "
            "Install dotnet or set DOCUMENT_TOOL_DOTNET."
        )
        return None

    if os.path.exists(tool_path) or shutil.which(tool_path):
        return [tool_path]

    logger.warning("Document tool executable was not found: %s", tool_path)
    return None


def _subprocess_env_with_wsl_bridge(*variable_names: str) -> dict[str, str]:
    env = os.environ.copy()
    if not variable_names:
        return env

    existing_entries = [
        entry for entry in env.get("WSLENV", "").split(":") if entry
    ]
    existing_names = {entry.split("/", 1)[0] for entry in existing_entries}

    for variable_name in variable_names:
        if variable_name in env and variable_name not in existing_names:
            existing_entries.append(variable_name)
            existing_names.add(variable_name)

    if existing_entries:
        env["WSLENV"] = ":".join(existing_entries)

    return env


def convert_docx_to_pdf_with_syncfusion(docx_path: str, output_path: str | None = None) -> str | None:
    command = _resolve_tool_command()
    if not command:
        return None

    if not os.getenv("SYNCFUSION_LICENSE_KEY"):
        logger.warning("SYNCFUSION_LICENSE_KEY is not set; cannot run Syncfusion DOCX to PDF conversion.")
        return None

    output_pdf_path = output_path or str(Path(docx_path).with_suffix(".pdf"))
    if os.path.exists(output_pdf_path):
        logger.info("Removing stale intermediate PDF before Syncfusion regeneration: %s", output_pdf_path)
        os.remove(output_pdf_path)

    dotnet_path = command[0]
    process = subprocess.run(
        command
        + [
            "convert-docx-to-pdf",
            "--docx",
            _to_dotnet_path(docx_path, dotnet_path),
            "--output",
            _to_dotnet_path(output_pdf_path, dotnet_path),
        ],
        capture_output=True,
        text=True,
        check=False,
        env=_subprocess_env_with_wsl_bridge("SYNCFUSION_LICENSE_KEY"),
    )

    if process.stdout.strip():
        logger.info("Document tool stdout:\n%s", process.stdout.strip())
    if process.stderr.strip():
        logger.warning("Document tool stderr:\n%s", process.stderr.strip())

    if process.returncode != 0:
        logger.warning(
            "Document tool exited with code %s while converting %s",
            process.returncode,
            docx_path,
        )
        return None

    if not os.path.exists(output_pdf_path):
        logger.warning("Document tool reported success but did not create %s", output_pdf_path)
        return None

    return output_pdf_path


def _write_page_market_map(page_market_mapping: dict | None) -> str | None:
    if not page_market_mapping:
        return None

    normalized_map = {}
    for page_num, page_data in page_market_mapping.items():
        if isinstance(page_data, tuple):
            market_text = " ".join(str(part or "") for part in page_data)
        else:
            market_text = str(page_data or "")
        if market_text.strip():
            normalized_map[str(page_num)] = market_text.strip()

    if not normalized_map:
        return None

    fd, path = tempfile.mkstemp(prefix="matrix_page_markets_", suffix=".json")
    with os.fdopen(fd, "w", encoding="utf-8") as handle:
        json.dump(normalized_map, handle)
    return path


def rewrite_matrix_amounts_with_dotnet(docx_path: str, page_market_mapping: dict | None = None) -> bool:
    command = _resolve_tool_command()
    if not command:
        return False

    dotnet_path = command[0]
    page_market_map_path = _write_page_market_map(page_market_mapping)
    command_args = command + [
        "rewrite-matrix-amounts",
        "--docx",
        _to_dotnet_path(docx_path, dotnet_path),
    ]
    if page_market_map_path:
        command_args.extend(
            [
                "--page-markets-json",
                _to_dotnet_path(page_market_map_path, dotnet_path),
            ]
        )

    try:
        process = subprocess.run(
            command_args,
            capture_output=True,
            text=True,
            check=False,
            env=_subprocess_env_with_wsl_bridge(),
        )
    finally:
        if page_market_map_path and os.path.exists(page_market_map_path):
            os.remove(page_market_map_path)

    if process.stdout.strip():
        logger.info("Document tool stdout:\n%s", process.stdout.strip())
    if process.stderr.strip():
        logger.warning("Document tool stderr:\n%s", process.stderr.strip())

    if process.returncode != 0:
        logger.warning(
            "Document tool exited with code %s while rewriting Matrix amounts in %s",
            process.returncode,
            docx_path,
        )
        return False

    return True
