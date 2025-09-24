#!/usr/bin/env python3
"""Populate Google Sheets with CPF/CNPJ per grupo/cota."""

from __future__ import annotations

import argparse
import asyncio
import logging
import re
from typing import List, Optional, Tuple

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
    return f"'{escaped}'!{range_clause}"


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
        sheet_range: str,
        header_title: str,
        force: bool,
        delay: float,
    ) -> None:
        self.processor = FinalWorkingProcessor(config_path)
        self.sheet_name, self.data_range = parse_sheet_range(sheet_range)
        self.header_title = header_title
        self.normalized_header = normalize_header(header_title)
        self.force = force
        self.delay = delay

        if not self.processor.google_sheets_client:
            raise RuntimeError("Google Sheets ingestion is not enabled in config.yaml")
        self.sheets = self.processor.google_sheets_client

        self.grupo_index: Optional[int] = None
        self.cota_index: Optional[int] = None
        self.cpf_index: Optional[int] = None

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
        header_row = self._ensure_header()
        rows = self._load_rows(header_row)
        if not rows:
            LOGGER.warning("No data rows found to process")
            return

        if self.grupo_index is None or self.cota_index is None or self.cpf_index is None:
            raise RuntimeError("Header indices were not initialized correctly")

        column_values: List[str] = []
        total_rows = len(rows)
        processed = 0
        skipped_existing = 0
        filled = 0

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
                        column_values.append(existing_cpf)
                        LOGGER.warning(
                            "Skipping row %s due to missing grupo/cota (raw values: %s / %s)",
                            idx,
                            grupo_raw,
                            cota_raw,
                        )
                        continue

                    if existing_cpf and not self.force:
                        column_values.append(existing_cpf)
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

                    column_values.append((cpf_cnpj or "").strip())
                    if self.delay:
                        await asyncio.sleep(self.delay)

            finally:
                await context.close()
                await browser.close()

        if len(column_values) != total_rows:
            raise RuntimeError(
                "Mismatch between processed rows and collected values: %s vs %s"
                % (len(column_values), total_rows)
            )

        target_column_letter = column_index_to_letter(self.cpf_index + 1)
        target_range = format_range(
            self.sheet_name,
            f"{target_column_letter}2:{target_column_letter}{total_rows + 1}",
        )
        values_payload = [[value] for value in column_values]
        update_success = self.sheets.update_values(target_range, values_payload)
        if not update_success:
            raise RuntimeError("Failed to update CPF column in Google Sheets")

        LOGGER.info(
            "CPF population completed: %s rows processed, %s existing skipped, %s new values written",
            processed,
            skipped_existing,
            filled,
        )


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Populate Google Sheets with CPF/CNPJ data.")
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to configuration YAML file (default: config.yaml)",
    )
    parser.add_argument(
        "--sheet-range",
        default="Página1!A:D",
        help="Sheet range containing GRUPO/COTA columns (default: Página1!A:D)",
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
    populator = CPFPopulator(
        config_path=args.config,
        sheet_range=args.sheet_range,
        header_title=args.header_title,
        force=args.force,
        delay=args.delay,
    )
    await populator.populate()


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()
    asyncio.run(run_async(args))


if __name__ == "__main__":
    main()
