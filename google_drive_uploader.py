"""Google Drive upload helper for boleto automation."""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional, Tuple

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload


class GoogleDriveUploader:
    """Handles uploads of boleto PDFs to Google Drive with year/month folders."""

    def __init__(
        self,
        credentials_path: str,
        drive_id: str,
        use_year_month_folders: bool = True,
        delegated_subject: Optional[str] = None,
        base_path: Optional[Path] = None,
        logger: Optional[logging.Logger] = None,
    ) -> None:
        self.logger = logger or logging.getLogger(__name__)

        self.base_path = Path(base_path) if base_path else Path.cwd()
        self.credentials_path = self._resolve_credentials_path(credentials_path)
        self.drive_id = drive_id
        self.use_year_month_folders = use_year_month_folders
        self.delegated_subject = delegated_subject

        self.service = None
        self.enabled = bool(self.credentials_path and self.drive_id)
        self._folder_cache: Dict[Tuple[str, str], str] = {}
        self._root_drive_id: Optional[str] = None
        self._is_shared_drive = False
        self.disabled_reason: Optional[str] = None

        if self.enabled:
            self._build_service()

    def _resolve_credentials_path(self, credentials_path: Optional[str]) -> Optional[Path]:
        if not credentials_path:
            return None

        path = Path(credentials_path).expanduser()
        if not path.is_absolute():
            path = self.base_path / path
        return path

    def _build_service(self) -> None:
        if not self.credentials_path or not self.credentials_path.exists():
            self.logger.error(
                "Google Drive credentials file not found: %s", self.credentials_path
            )
            self._disable("Google Drive credentials file not found")
            return

        try:
            scopes = ["https://www.googleapis.com/auth/drive"]
            credentials = service_account.Credentials.from_service_account_file(
                str(self.credentials_path), scopes=scopes
            )
            if self.delegated_subject:
                credentials = credentials.with_subject(self.delegated_subject)

            self.service = build(
                "drive",
                "v3",
                credentials=credentials,
                cache_discovery=False,
            )

            self.enabled = True
            self.disabled_reason = None
            self._inspect_root_folder()
        except Exception as error:  # Broad catch to avoid crashing automation
            self.logger.error("Failed to initialize Google Drive service: %s", error)
            self._disable("Failed to initialize Google Drive service")

    def _inspect_root_folder(self) -> None:
        if not self.service or not self.drive_id:
            return

        try:
            metadata = (
                self.service.files()
                .get(
                    fileId=self.drive_id,
                    fields="id, name, mimeType, driveId",
                    supportsAllDrives=True,
                )
                .execute()
            )

            if metadata.get("mimeType") != "application/vnd.google-apps.folder":
                self.logger.error(
                    "Configured Google Drive ID %s is not a folder (mimeType=%s)",
                    self.drive_id,
                    metadata.get("mimeType"),
                )
                self._disable("Configured drive_id is not a folder")
                return

            self._root_drive_id = metadata.get("driveId")
            self._is_shared_drive = bool(self._root_drive_id and self._root_drive_id != "root")

            self.logger.info(
                "Google Drive uploader ready (folder='%s', shared_drive=%s)",
                metadata.get("name"),
                self._is_shared_drive,
            )
        except HttpError as error:
            self.logger.error(
                "Unable to access Google Drive folder %s: %s",
                self.drive_id,
                error,
            )
            self._disable("Unable to access configured drive folder")

    def upload_pdf(
        self,
        local_path: str,
        file_name: str,
        reference_date: Optional[datetime] = None,
    ) -> Optional[str]:
        if not self.enabled or not self.service:
            if self.disabled_reason:
                self.logger.debug(
                    "Skipping Google Drive upload because uploader is disabled: %s",
                    self.disabled_reason,
                )
            return None

        file_path = Path(local_path)
        if not file_path.exists():
            self.logger.error("Local file not found for upload: %s", local_path)
            return None

        ref_date = reference_date or datetime.now()
        parent_id = self.drive_id

        if self.use_year_month_folders:
            year_folder = self._get_or_create_folder(ref_date.strftime("%Y"), parent_id)
            if not year_folder:
                return None
            parent_id = year_folder

            month_folder = self._get_or_create_folder(ref_date.strftime("%m"), parent_id)
            if not month_folder:
                return None
            parent_id = month_folder

        try:
            metadata = {
                "name": file_name,
                "parents": [parent_id],
                "mimeType": "application/pdf",
            }
            media = MediaFileUpload(str(file_path), mimetype="application/pdf", resumable=False)
            created_file = (
                self.service.files()
                .create(
                    body=metadata,
                    media_body=media,
                    fields="id, webViewLink",
                    supportsAllDrives=True,
                )
                .execute()
            )
            file_id = created_file.get("id")
            web_link = created_file.get("webViewLink")
            self.logger.info(
                "Uploaded %s to Google Drive (file_id=%s, link=%s)",
                file_name,
                file_id,
                web_link,
            )
            return file_id
        except HttpError as error:
            reason = ""
            try:
                reason = error._get_reason()  # type: ignore[attr-defined]
            except Exception:
                reason = ""
            error_text = str(error)
            reason_lower = reason.lower()
            error_text_lower = error_text.lower()
            self.logger.error("Google Drive upload failed for %s: %s", file_name, error_text)

            if getattr(error.resp, "status", None) == 403 and (
                "storagequota" in reason_lower
                or "storage quota" in reason_lower
                or "storagequotaexceeded" in error_text_lower
            ):
                self.logger.warning(
                    "Google Drive reported storage quota exceeded; disabling uploader for this run"
                )
                self._disable("Google Drive storage quota exceeded")
        except Exception as error:  # Catch-all to keep automation running
            self.logger.error("Unexpected error during Google Drive upload: %s", error)
        return None

    def _get_or_create_folder(self, folder_name: str, parent_id: str) -> Optional[str]:
        cache_key = (parent_id, folder_name)
        if cache_key in self._folder_cache:
            return self._folder_cache[cache_key]

        try:
            query = (
                f"name = '{folder_name}' "
                "and mimeType = 'application/vnd.google-apps.folder' "
                "and trashed = false "
                f"and '{parent_id}' in parents"
            )
            list_kwargs = {
                "q": query,
                "spaces": "drive",
                "fields": "files(id, name)",
                "includeItemsFromAllDrives": True,
                "supportsAllDrives": True,
            }
            if self._is_shared_drive and self._root_drive_id:
                list_kwargs["driveId"] = self._root_drive_id
                list_kwargs["corpora"] = "drive"

            response = self.service.files().list(**list_kwargs).execute()
            folders = response.get("files", [])
            if folders:
                folder_id = folders[0]["id"]
                self._folder_cache[cache_key] = folder_id
                return folder_id

            metadata = {
                "name": folder_name,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [parent_id],
            }
            folder = (
                self.service.files()
                .create(body=metadata, fields="id", supportsAllDrives=True)
                .execute()
            )
            folder_id = folder.get("id")
            self._folder_cache[cache_key] = folder_id
            self.logger.info(
                "Created Google Drive folder %s under parent %s (id=%s)",
                folder_name,
                parent_id,
                folder_id,
            )
            return folder_id
        except HttpError as error:
            self.logger.error(
                "Failed to locate or create Google Drive folder '%s': %s",
                folder_name,
                error,
            )
        except Exception as error:
            self.logger.error(
                "Unexpected error while ensuring Google Drive folder '%s': %s",
                folder_name,
                error,
            )
        return None

    def _disable(self, reason: str) -> None:
        if not self.enabled and self.disabled_reason == reason:
            return
        if self.service is not None:
            self.service = None
        if self.enabled:
            self.logger.warning("Disabling Google Drive uploader: %s", reason)
        self.enabled = False
        self.disabled_reason = reason
