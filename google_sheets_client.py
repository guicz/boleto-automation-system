"""Google Sheets helper for boleto automation."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Optional

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class GoogleSheetsClient:
    """Fetches records from a Google Sheets spreadsheet."""

    READ_ONLY_SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    READ_WRITE_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    def __init__(
        self,
        credentials_path: Path,
        spreadsheet_id: str,
        logger: logging.Logger,
        scopes: Optional[List[str]] = None,
    ) -> None:
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.logger = logger
        self.service = None
        self.scopes = scopes or self.READ_ONLY_SCOPES

    def _get_service(self):
        if self.service:
            return self.service
        credentials = service_account.Credentials.from_service_account_file(
            str(self.credentials_path), scopes=self.scopes
        )
        self.service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
        return self.service

    def fetch_records(self, sheet_range: str) -> List[Dict[str, str]]:
        try:
            service = self._get_service()
            sheet = service.spreadsheets()
            result = (
                sheet.values()
                .get(spreadsheetId=self.spreadsheet_id, range=sheet_range)
                .execute()
            )
            values = result.get("values", [])
            if not values:
                self.logger.warning("Google Sheets returned no rows for range %s", sheet_range)
                return []

            headers = [header.strip().upper() for header in values[0]]
            records = []
            for row in values[1:]:
                record = {header: (row[idx] if idx < len(row) else "") for idx, header in enumerate(headers)}
                records.append(record)
            self.logger.info("Fetched %d records from Google Sheets", len(records))
            return records
        except HttpError as error:
            self.logger.error("Google Sheets API error: %s", error)
        except Exception as error:  # pragma: no cover - defensive
            self.logger.error("Unexpected error fetching Google Sheets data: %s", error)
        return []

    def get_values(self, sheet_range: str) -> List[List[str]]:
        try:
            service = self._get_service()
            result = (
                service.spreadsheets()
                .values()
                .get(spreadsheetId=self.spreadsheet_id, range=sheet_range)
                .execute()
            )
            values = result.get("values", [])
            return values
        except HttpError as error:
            self.logger.error("Google Sheets get_values error: %s", error)
        except Exception as error:  # pragma: no cover - defensive
            self.logger.error("Unexpected error getting Google Sheets values: %s", error)
        return []

    def update_values(self, sheet_range: str, values: List[List[str]]) -> bool:
        try:
            service = self._get_service()
            body = {"values": values}
            response = (
                service.spreadsheets()
                .values()
                .update(
                    spreadsheetId=self.spreadsheet_id,
                    range=sheet_range,
                    valueInputOption="RAW",
                    body=body,
                )
                .execute()
            )
            updated_cells = response.get("updatedCells", 0)
            if updated_cells:
                self.logger.debug(
                    "Updated %s cells in range %s",
                    updated_cells,
                    sheet_range,
                )
                return True
            self.logger.warning("No cells were updated for range %s", sheet_range)
        except HttpError as error:
            self.logger.error("Google Sheets update error: %s", error)
        except Exception as error:  # pragma: no cover - defensive
            self.logger.error("Unexpected error updating Google Sheets: %s", error)
        return False

    def append_row(self, sheet_range: str, values: List[str]) -> bool:
        try:
            service = self._get_service()
            body = {"values": [values]}
            response = (
                service.spreadsheets()
                .values()
                .append(
                    spreadsheetId=self.spreadsheet_id,
                    range=sheet_range,
                    valueInputOption="USER_ENTERED",
                    insertDataOption="INSERT_ROWS",
                    body=body,
                )
                .execute()
            )
            updates = response.get("updates", {})
            updated_rows = updates.get("updatedRows", 0)
            if updated_rows:
                self.logger.debug("Appended row to Google Sheets (%s)", sheet_range)
                return True
            self.logger.warning("No rows appended to Google Sheets for range %s", sheet_range)
        except HttpError as error:
            self.logger.error("Google Sheets append error: %s", error)
        except Exception as error:
            self.logger.error("Unexpected error appending to Google Sheets: %s", error)
        return False
