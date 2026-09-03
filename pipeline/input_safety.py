"""Dependency-free input budgets shared by the CLI and cloud pipeline."""

import zipfile
from pathlib import Path


MAX_PPTX_BYTES = 50 * 1024 * 1024
MAX_ZIP_MEMBERS = 5000
MAX_ZIP_UNCOMPRESSED_BYTES = 600 * 1024 * 1024
MAX_ZIP_MEMBER_BYTES = 200 * 1024 * 1024
MAX_XML_MEMBER_BYTES = 20 * 1024 * 1024
MAX_ZIP_EXPANSION_RATIO = 150
MAX_SLIDES = 100
MAX_OFFENSE_PLAYS = 64
MAX_DEFENSE_PLAYS = 24


def validate_pptx_archive(pptx_path):
    """Reject malformed or unexpectedly expensive PPTX archives up front."""
    path = Path(pptx_path)
    if not path.is_file():
        raise FileNotFoundError(f"PPTX file not found: {path}")
    if path.suffix.lower() != ".pptx":
        raise ValueError("Input must be a .pptx file")
    if path.stat().st_size <= 0:
        raise ValueError("PPTX file is empty")
    if path.stat().st_size > MAX_PPTX_BYTES:
        raise ValueError("PPTX file is too large (max 50 MB)")

    try:
        with zipfile.ZipFile(path, "r") as archive:
            infos = archive.infolist()
            if not infos or len(infos) > MAX_ZIP_MEMBERS:
                raise ValueError(f"PPTX archive has too many entries (max {MAX_ZIP_MEMBERS})")

            names = {info.filename for info in infos}
            required = {"[Content_Types].xml", "ppt/presentation.xml"}
            if not required.issubset(names):
                raise ValueError("File is not a valid PowerPoint presentation")

            total_uncompressed = 0
            total_compressed = 0
            normalized_names = set()
            for info in infos:
                normalized = info.filename.replace("\\", "/")
                parts = [part for part in normalized.split("/") if part]
                if normalized.startswith("/") or ".." in parts:
                    raise ValueError("PPTX archive contains an unsafe path")
                if normalized in normalized_names:
                    raise ValueError("PPTX archive contains duplicate entries")
                normalized_names.add(normalized)
                if info.flag_bits & 0x1:
                    raise ValueError("Encrypted PPTX archives are not supported")
                if info.file_size > MAX_ZIP_MEMBER_BYTES:
                    raise ValueError("PPTX archive contains an oversized entry")
                if normalized.lower().endswith((".xml", ".rels")) and info.file_size > MAX_XML_MEMBER_BYTES:
                    raise ValueError("PPTX archive contains an oversized XML entry")
                total_uncompressed += info.file_size
                total_compressed += info.compress_size

            if total_uncompressed > MAX_ZIP_UNCOMPRESSED_BYTES:
                raise ValueError("PPTX expands beyond the 600 MB safety limit")
            ratio = total_uncompressed / max(1, total_compressed)
            if ratio > MAX_ZIP_EXPANSION_RATIO:
                raise ValueError("PPTX archive has a suspicious compression ratio")
    except zipfile.BadZipFile as exc:
        raise ValueError("File is not a valid PPTX archive") from exc


def validate_print_play_counts(plays):
    offense = sum(1 for play in plays if play["section"] == "OFFENSE")
    defense = sum(1 for play in plays if play["section"] == "DEFENSE")
    if offense > MAX_OFFENSE_PLAYS:
        raise ValueError(
            f"Playbook has {offense} offense plays; printed outputs support at most {MAX_OFFENSE_PLAYS}"
        )
    if defense > MAX_DEFENSE_PLAYS:
        raise ValueError(
            f"Playbook has {defense} defense plays; printed outputs support at most {MAX_DEFENSE_PLAYS}"
        )
