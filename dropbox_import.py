from __future__ import annotations

import hashlib
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional

import dropbox
from dropbox.files import FileMetadata


@dataclass
class DropboxCSVFile:
    name: str
    path_lower: str
    server_modified: str
    size: int
    content_hash: str
    local_path: Path


def get_dropbox_client(access_token: str) -> dropbox.Dropbox:
    if not access_token:
        raise ValueError("Missing Dropbox access token.")
    return dropbox.Dropbox(access_token, timeout=60)


def list_csv_files(dbx: dropbox.Dropbox, folder_path: str) -> list[FileMetadata]:
    folder_path = (folder_path or "").strip() or ""
    entries: list[FileMetadata] = []
    result = dbx.files_list_folder(folder_path, recursive=False)
    while True:
        for entry in result.entries:
            if isinstance(entry, FileMetadata) and entry.name.lower().endswith(".csv"):
                entries.append(entry)
        if not result.has_more:
            break
        result = dbx.files_list_folder_continue(result.cursor)
    entries.sort(key=lambda e: (e.server_modified, e.path_lower or e.name))
    return entries


def download_file(dbx: dropbox.Dropbox, metadata: FileMetadata, download_dir: Path) -> DropboxCSVFile:
    download_dir.mkdir(parents=True, exist_ok=True)
    safe_name = metadata.name.replace("/", "_").replace("\\", "_")
    local_path = download_dir / safe_name
    _, response = dbx.files_download(metadata.path_lower)
    local_path.write_bytes(response.content)
    return DropboxCSVFile(
        name=metadata.name,
        path_lower=metadata.path_lower or metadata.name,
        server_modified=metadata.server_modified.isoformat() if metadata.server_modified else "",
        size=int(metadata.size or 0),
        content_hash=metadata.content_hash or hashlib.sha256(response.content).hexdigest(),
        local_path=local_path,
    )


def download_new_csvs(access_token: str, folder_path: str, download_dir: Path) -> list[DropboxCSVFile]:
    dbx = get_dropbox_client(access_token)
    files = list_csv_files(dbx, folder_path)
    return [download_file(dbx, f, download_dir) for f in files]
