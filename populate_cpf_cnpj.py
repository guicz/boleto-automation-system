#!/usr/bin/env python3
"""Populate Google Sheets with CPF/CNPJ per grupo/cota."""

from __future__ import annotations

import argparse
import asyncio
import logging
import re
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
from playwright.async_api import async_playwright

from final_working_boleto_processor import FinalWorkingProcessor


LOGGER = logging.getLogger("cpf_populator")


def column_index_to_letter(index: int) -> str:
    if index < 1:
        raise ValueError("Column index must be 1 or greater")
    letters: List[str] = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def normalize_header(value: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", (value or "").strip().upper())


def sanitize_sheet_name(name: str) -> str:
    name = name.strip()
    if name.startswith("'") and name.endswith("'") and len(name) >= 2:
        name = name[1:-1]
    return name


def format_range(sheet: str, range_clause: str) -> str:
    escaped = sheet.replace("'", "''")
    needs_quotes = not re.fullmatch(r"\w+", sheet)
    if needs_quotes:
        return f"'{escaped}'!{range_clause}"
    return f"{sheet}!{range_clause}"


def parse_sheet_range(value: str) -> Tuple[str, str]:
    if "!" in value:
        sheet, rng = value.split("!", 1)
    else:
        sheet, rng = value, "A:Z"
    return sanitize_sheet_name(sheet), rng


class CPFPopulator:
    def __init__(
        self,
        config_path: str,
        header_title: str,
        force: bool,
        delay: float,
        sheet_range: Optional[str] = None,
        csv_path: Optional[str] = None,
        csv_encoding: str = "utf-8",
        csv_delimiter: str = ",",
        flush_every: int = 1,
    ) -> None:
        if not sheet_range and not csv_path:
            raise ValueError("Either sheet_range or csv_path must be provided")
        if sheet_range and csv_path:
            raise ValueError("Provide only one of sheet_range or csv_path")

        self.processor = FinalWorkingProcessor(config_path)
        self.header_title = header_title
        self.normalized_header = normalize_header(header_title)
        self.force = force
        self.delay = delay

        self.sheet_range = sheet_range
        self.csv_path = Path(csv_path).resolve() if csv_path else None
        self.csv_encoding = csv_encoding
        self.csv_delimiter = csv_delimiter
        self.flush_every = max(1, flush_every)

        self.sheets = None
        self.sheet_name: Optional[str] = None
        self.data_range: Optional[str] = None
        self.grupo_index: Optional[int] = None
        self.cota_index: Optional[int] = None
        self.cpf_index: Optional[int] = None

        if self.sheet_range:
            if not self.processor.google_sheets_client:
                raise RuntimeError("Google Sheets ingestion is not enabled in config.yaml")
            self.sheets = self.processor.google_sheets_client
            self.sheet_name, self.data_range = parse_sheet_range(self.sheet_range)

    def _ensure_header(self) -> List[str]:
        header_values = self.sheets.get_values(format_range(self.sheet_name, "1:1"))
        if not header_values or not header_values[0]:
            raise RuntimeError(
                f"Worksheet {self.sheet_name} appears to have an empty header row"
            )

        header_row = header_values[0]
        normalized_headers = [normalize_header(h) for h in header_row]

        if self.normalized_header not in normalized_headers:
            header_row.append(self.header_title)
            success = self.sheets.update_values(
                f"{self.sheet_name}!1:1", [header_row]
            )
            if not success:
                raise RuntimeError("Failed to append CPF header to worksheet")
            normalized_headers.append(self.normalized_header)
            LOGGER.info("Added new header column '%s'", self.header_title)

        try:
            self.grupo_index = normalized_headers.index("GRUPO")
            self.cota_index = normalized_headers.index("COTA")
        except ValueError as error:  # pragma: no cover - configuration issue
            raise RuntimeError("Worksheet must contain GRUPO and COTA columns") from error

        self.cpf_index = normalized_headers.index(self.normalized_header)
        return header_row

    def _load_rows(self, header_row: List[str]) -> List[List[str]]:
        column_count = len(header_row)
        end_column = column_index_to_letter(column_count)
        data_range = format_range(self.sheet_name, f"A2:{end_column}")
        rows = self.sheets.get_values(data_range)
        return rows

    async def populate(self) -> None:
        if self.csv_path:
            await self._populate_csv()
        else:
            await self._populate_sheet()

    async def _populate_sheet(self) -> None:
        header_row = self._ensure_header()
        rows = self._load_rows(header_row)
        if self.grupo_index is None or self.cota_index is None or self.cpf_index is None:
            raise RuntimeError("Header indices were not initialized correctly")

        if not rows:
            LOGGER.warning("No data rows found to process")
            return

        processed = 0
        skipped_existing = 0
        filled = 0
        pending_updates: List[Tuple[int, str]] = []

        browser_config = self.processor.config.get("browser", {})
        headless = browser_config.get("headless", True)
        slow_mo = browser_config.get("slow_mo")
        viewport = browser_config.get("viewport", {"width": 1280, "height": 720})

        browser_args = [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--disable-web-security",
            "--disable-background-timer-throttling",
            "--disable-renderer-backgrounding",
            "--disable-popup-blocking",
            "--print-to-pdf-no-header",
            "--run-all-compositor-stages-before-draw",
        ]

        async with async_playwright() as p:
            launch_kwargs = {"headless": headless, "args": browser_args}
            if slow_mo is not None:
                launch_kwargs["slow_mo"] = slow_mo

            browser = await p.chromium.launch(**launch_kwargs)

            context_kwargs = {"accept_downloads": False}
            if viewport:
                context_kwargs["viewport"] = viewport

            context = await browser.new_context(**context_kwargs)
            page = await context.new_page()

            try:
                login_success = await self.processor.login(page)
                if not login_success:
                    raise RuntimeError("Login failed; cannot continue CPF population")

                for idx, row in enumerate(rows, start=2):
                    grupo_raw = row[self.grupo_index] if self.grupo_index < len(row) else ""
                    cota_raw = row[self.cota_index] if self.cota_index < len(row) else ""
                    existing_cpf_raw = row[self.cpf_index] if self.cpf_index < len(row) else ""
                    existing_cpf = existing_cpf_raw.strip()

                    grupo = self.processor.sanitize_grupo(grupo_raw)
                    cota = self.processor.sanitize_cota(cota_raw)

                    if not grupo or not cota:
                        LOGGER.warning(
                            "Skipping row %s due to missing grupo/cota (raw values: %s / %s)",
                            idx,
                            grupo_raw,
                            cota_raw,
                        )
                        continue

                    if existing_cpf and not self.force:
                        skipped_existing += 1
                        continue

                    processed += 1
                    success, search_result = await self.processor.search_record(page, grupo, cota)
                    cpf_cnpj = existing_cpf
                    if success:
                        cpf_cnpj = search_result.get("cpf_cnpj") or existing_cpf
                        if cpf_cnpj:
                            filled += 1 if not existing_cpf else 0
                        status = search_result.get("contemplado_status", "")
                        LOGGER.info(
                            "Row %s (%s/%s) → CPF=%s Status=%s",
                            idx,
                            grupo,
                            cota,
                            cpf_cnpj or "",
                            status,
                        )
                    else:
                        LOGGER.error(
                            "Unable to fetch CPF for %s/%s (row %s): %s",
                            grupo,
                            cota,
                            idx,
                            search_result.get("error"),
                        )

                    new_value = (cpf_cnpj or "").strip()
                    if new_value and new_value != existing_cpf:
                        pending_updates.append((idx, new_value))
                        if len(pending_updates) >= self.flush_every:
                            self._flush_sheet_updates(pending_updates)
                            pending_updates.clear()
                    elif self.force:
                        pending_updates.append((idx, new_value))
                        if len(pending_updates) >= self.flush_every:
                            self._flush_sheet_updates(pending_updates)
                            pending_updates.clear()
                    if self.delay:
                        await asyncio.sleep(self.delay)

            finally:
                await context.close()
                await browser.close()

        if pending_updates:
            self._flush_sheet_updates(pending_updates)

        if processed == 0 and skipped_existing == 0:
            LOGGER.info("No updates were necessary for the selected sheet range")

        LOGGER.info(
            "CPF population completed: %s rows processed, %s existing skipped, %s new values written",
            processed,
            skipped_existing,
            filled,
        )

    def _flush_sheet_updates(self, updates: List[Tuple[int, str]]) -> None:
        if not updates:
            return
        if not self.sheets:
            raise RuntimeError("Sheets client not configured")
        target_column_letter = column_index_to_letter(self.cpf_index + 1)
        for row_index, value in updates:
            target = format_range(self.sheet_name, f"{target_column_letter}{row_index}")
            success = self.sheets.update_values(target, [[value]])
            if not success:
                raise RuntimeError(f"Failed to update cell {target}")
        LOGGER.debug("Flushed %s sheet updates", len(updates))

    async def _populate_csv(self) -> None:
        if not self.csv_path:
            raise RuntimeError("CSV path not configured")

        if not self.csv_path.exists():
            raise FileNotFoundError(f"CSV file not found: {self.csv_path}")

        df = pd.read_csv(
            self.csv_path,
            sep=self.csv_delimiter,
            encoding=self.csv_encoding,
            dtype=str,
        ).fillna("")

        headers = list(df.columns)
        if "GRUPO" not in headers or "COTA" not in headers:
            raise RuntimeError("CSV must contain GRUPO and COTA columns")

        if self.header_title not in headers:
            df[self.header_title] = ""
            headers.append(self.header_title)
            LOGGER.info("Added new column '%s' to CSV", self.header_title)

        processed = 0
        skipped_existing = 0
        filled = 0
        pending_dirty = 0

        browser_config = self.processor.config.get("browser", {})
        headless = browser_config.get("headless", True)
        slow_mo = browser_config.get("slow_mo")
        viewport = browser_config.get("viewport", {"width": 1280, "height": 720})

        browser_args = [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--disable-web-security",
            "--disable-background-timer-throttling",
            "--disable-renderer-backgrounding",
            "--disable-popup-blocking",
            "--print-to-pdf-no-header",
            "--run-all-compositor-stages-before-draw",
        ]

        async with async_playwright() as p:
            launch_kwargs = {"headless": headless, "args": browser_args}
            if slow_mo is not None:
                launch_kwargs["slow_mo"] = slow_mo

            browser = await p.chromium.launch(**launch_kwargs)

            context_kwargs = {"accept_downloads": False}
            if viewport:
                context_kwargs["viewport"] = viewport

            context = await browser.new_context(**context_kwargs)
            page = await context.new_page()

            try:
                login_success = await self.processor.login(page)
                if not login_success:
                    raise RuntimeError("Login failed; cannot continue CPF population")

                for idx, row in df.iterrows():
                    grupo_raw = row.get("GRUPO", "")
                    cota_raw = row.get("COTA", "")
                    existing_doc = str(row.get(self.header_title, "")).strip()

                    grupo = self.processor.sanitize_grupo(grupo_raw)
                    cota = self.processor.sanitize_cota(cota_raw)

                    if not grupo or not cota:
                        LOGGER.warning(
                            "Skipping CSV row %s due to missing grupo/cota (raw: %s / %s)",
                            idx + 2,
                            grupo_raw,
                            cota_raw,
                        )
                        continue

                    if existing_doc and not self.force:
                        skipped_existing += 1
                        continue

                    processed += 1
                    success, search_result = await self.processor.search_record(page, grupo, cota)
                    cpf_cnpj = existing_doc
                    if success:
                        cpf_cnpj = search_result.get("cpf_cnpj") or existing_doc
                        status = search_result.get("contemplado_status", "")
                        LOGGER.info(
                            "CSV row %s (%s/%s) → CPF=%s Status=%s",
                            idx + 2,
                            grupo,
                            cota,
                            cpf_cnpj or "",
                            status,
                        )
                        if cpf_cnpj and not existing_doc:
                            filled += 1
                    else:
                        LOGGER.error(
                            "Unable to fetch CPF for %s/%s (CSV row %s): %s",
                            grupo,
                            cota,
                            idx + 2,
                            search_result.get("error"),
                        )

                    new_value = (cpf_cnpj or "").strip()
                    if new_value != existing_doc or (self.force and new_value):
                        df.at[idx, self.header_title] = new_value
                        pending_dirty += 1
                        if new_value and not existing_doc:
                            filled += 1
                    if pending_dirty >= self.flush_every:
                        self._flush_csv(df)
                        pending_dirty = 0
                    if self.delay:
                        await asyncio.sleep(self.delay)

            finally:
                await context.close()
                await browser.close()

        if pending_dirty:
            self._flush_csv(df)

        LOGGER.info(
            "CSV population completed: %s rows processed, %s existing skipped, %s new values written",
            processed,
            skipped_existing,
            filled,
        )

    def _flush_csv(self, df: pd.DataFrame) -> None:
        if not self.csv_path:
            return
        df.to_csv(
            self.csv_path,
            sep=self.csv_delimiter,
            encoding=self.csv_encoding,
            index=False,
        )
        LOGGER.debug("CSV file flushed to disk (%s)", self.csv_path)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Populate CPF/CNPJ data into Google Sheets or a local CSV."
    )
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to configuration YAML file (default: config.yaml)",
    )
    parser.add_argument(
        "--sheet-range",
        help="Sheet range (e.g. 'Página1!A:D') to update in Google Sheets",
    )
    parser.add_argument(
        "--csv-path",
        help="Local CSV file to update (mutually exclusive with --sheet-range)",
    )
    parser.add_argument(
        "--csv-delimiter",
        default=",",
        help="Delimiter used in the CSV (default: ',')",
    )
    parser.add_argument(
        "--csv-encoding",
        default="utf-8",
        help="CSV encoding (default: utf-8)",
    )
    parser.add_argument(
        "--flush-every",
        type=int,
        default=1,
        help="Persist changes every N updates (default: 1)",
    )
    parser.add_argument(
        "--header-title",
        default="CPF/CNPJ",
        help="Header title for the CPF column (default: CPF/CNPJ)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Reprocess rows that already have CPF values",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.0,
        help="Delay in seconds between lookups (default: 0)",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging level (default: INFO)",
    )
    return parser


async def run_async(args: argparse.Namespace) -> None:
    logging.basicConfig(
        level=getattr(logging, args.log_level.upper()),
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    sheet_range = args.sheet_range
    csv_path = args.csv_path
    if not sheet_range and not csv_path:
        sheet_range = "Página1!A:D"
    populator = CPFPopulator(
        config_path=args.config,
        header_title=args.header_title,
        force=args.force,
        delay=args.delay,
        sheet_range=sheet_range,
        csv_path=csv_path,
        csv_encoding=args.csv_encoding,
        csv_delimiter=args.csv_delimiter,
        flush_every=args.flush_every,
    )
    await populator.populate()


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()
    asyncio.run(run_async(args))


if __name__ == "__main__":
    main()
