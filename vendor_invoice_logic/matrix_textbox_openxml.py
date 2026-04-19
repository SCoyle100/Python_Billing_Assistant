import json
import logging
import os
import shutil
import subprocess
import tempfile
from typing import Mapping


logger = logging.getLogger(__name__)


def _candidate_dotnet_commands() -> list[str]:
    candidates = []

    configured_dotnet = os.getenv("MATRIX_OPENXML_DOTNET")
    if configured_dotnet:
        candidates.append(configured_dotnet)

    path_dotnet = shutil.which("dotnet")
    if path_dotnet:
        candidates.append(path_dotnet)

    # WSL can usually invoke the Windows SDK directly. This lets a Windows-side
    # .NET install build/run the helper before Ubuntu has its own SDK.
    candidates.extend(
        [
            "/mnt/c/Program Files/dotnet/dotnet.exe",
            "/mnt/c/Program Files (x86)/dotnet/dotnet.exe",
        ]
    )

    return candidates


def _resolve_tool_command() -> list[str] | None:
    tool_path = os.getenv("MATRIX_OPENXML_TEXTBOX_TOOL")
    if not tool_path:
        return None

    if tool_path.endswith(".dll"):
        for dotnet_path in _candidate_dotnet_commands():
            if os.path.exists(dotnet_path) or shutil.which(dotnet_path):
                return [dotnet_path, tool_path]

        logger.warning(
            "MATRIX_OPENXML_TEXTBOX_TOOL points to a .dll, but no dotnet host was found. "
            "Install dotnet on Ubuntu or set MATRIX_OPENXML_DOTNET to a dotnet executable."
        )
        return None

    return [tool_path]


def rewrite_textboxes_with_openxml(
    docx_path: str,
    replacements: Mapping[str, str],
) -> bool:
    """
    Invoke the optional C# Open XML textbox rewriter.

    Returns True only when the external tool runs successfully.
    """
    command = _resolve_tool_command()
    if not command:
        return False

    replacement_items = [
        {"original": original, "replacement": replacement}
        for original, replacement in replacements.items()
        if original and replacement and original != replacement
    ]
    if not replacement_items:
        logger.info("No textbox replacements to send to the Open XML helper.")
        return False

    payload = {"replacements": replacement_items}

    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".json",
        delete=False,
        encoding="utf-8",
    ) as temp_file:
        json.dump(payload, temp_file, indent=2)
        replacements_path = temp_file.name

    try:
        process = subprocess.run(
            command
            + [
                "rewrite-textboxes",
                "--docx",
                docx_path,
                "--replacements",
                replacements_path,
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        if process.stdout.strip():
            logger.info("Open XML textbox tool stdout:\n%s", process.stdout.strip())
        if process.stderr.strip():
            logger.warning("Open XML textbox tool stderr:\n%s", process.stderr.strip())

        if process.returncode == 2:
            logger.info(
                "Open XML textbox tool found no matching textbox replacements for %s",
                docx_path,
            )
            return False

        if process.returncode != 0:
            logger.warning(
                "Open XML textbox tool exited with code %s for %s",
                process.returncode,
                docx_path,
            )
            return False

        logger.info("Open XML textbox tool rewrote textboxes for %s", docx_path)
        return True
    except OSError as exc:
        logger.warning("Failed to launch Open XML textbox tool: %s", exc)
        return False
    finally:
        try:
            os.unlink(replacements_path)
        except OSError:
            pass
