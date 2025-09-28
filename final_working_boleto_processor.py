#!/usr/bin/env python3
"""
Final Working Boleto Processor
SOLUTION: Parse the onclick attribute to construct and submit a POST request directly to Slip.asp, 
replicating the website's form submission for maximum reliability.
"""

import asyncio
import argparse
import ast
import io
import json
import math
import logging
import os
import sys
import re
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
import yaml
from playwright.async_api import async_playwright, Browser, Page, BrowserContext

from google_drive_uploader import GoogleDriveUploader
from google_sheets_client import GoogleSheetsClient
from notifier import WebhookNotifier
from file_link_service import FileLinkService


class ProcessedRecordTracker:
    """Persists successfully processed grupo/cota combinations to avoid duplicates."""

    def __init__(
        self,
        path: Path,
        retention_days: Optional[int],
        logger: logging.Logger,
    ) -> None:
        self.path = path
        self.retention_days = retention_days
        self.logger = logger
        self.records: Dict[str, Dict] = {}

        self._load()
        if self._purge_expired():
            self._save()

    def _load(self) -> None:
        if not self.path.exists():
            return

        try:
            with open(self.path, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
            if isinstance(data, dict):
                self.records = data
            else:
                self.logger.warning(
                    "Processed state file %s is invalid; starting fresh",
                    self.path,
                )
        except Exception as error:
            self.logger.error(
                "Failed to load processed state file %s: %s",
                self.path,
                error,
            )

    def _save(self) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            tmp_path = self.path.with_suffix(self.path.suffix + '.tmp')
            with open(tmp_path, 'w', encoding='utf-8') as handle:
                json.dump(self.records, handle, indent=2, ensure_ascii=False)
            tmp_path.replace(self.path)
        except Exception as error:
            self.logger.error("Failed to persist processed state to %s: %s", self.path, error)

    def _purge_expired(self) -> bool:
        if not self.retention_days:
            return False

        threshold = datetime.now() - timedelta(days=self.retention_days)
        purged = False
        for key, value in list(self.records.items()):
            timestamp = value.get('timestamp')
            if not timestamp:
                continue
            try:
                record_datetime = datetime.fromisoformat(timestamp)
            except ValueError:
                continue
            if record_datetime < threshold:
                del self.records[key]
                purged = True

        if purged:
            self.logger.info(
                "Purged processed records older than %s days",
                self.retention_days,
            )
        return purged

    def _make_key(self, grupo: str, cota: str) -> str:
        return f"{grupo.strip()}|{cota.strip()}"

    def is_processed(self, grupo: str, cota: str) -> bool:
        if not grupo or not cota:
            return False
        return self._make_key(grupo, cota) in self.records

    def mark_processed(self, grupo: str, cota: str, metadata: Optional[Dict] = None) -> None:
        if not grupo or not cota:
            return

        metadata = metadata or {}
        record = {
            'grupo': grupo,
            'cota': cota,
            'timestamp': metadata.get('timestamp') or datetime.now().isoformat(),
        }

        if 'downloaded_files' in metadata:
            record['downloaded_files'] = metadata['downloaded_files']
        if 'cpf_cnpj' in metadata and metadata['cpf_cnpj']:
            record['cpf_cnpj'] = metadata['cpf_cnpj']
        if 'drive_file_ids' in metadata and metadata['drive_file_ids']:
            record['drive_file_ids'] = metadata['drive_file_ids']

        self.records[self._make_key(grupo, cota)] = record
        self._save()



class ResumeManager:
    """Handles persistence of resume checkpoints for interrupted runs."""

    def __init__(self, path: Optional[Path], enabled: bool, logger: logging.Logger) -> None:
        self.path = path
        self.enabled = enabled and path is not None
        self.logger = logger

    def load_state(self) -> Dict[str, Dict[str, str]]:
        if not self.enabled or not self.path or not self.path.exists():
            return {}
        try:
            with open(self.path, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
            if isinstance(data, dict):
                return data
            self.logger.warning("Resume state file %s is invalid; ignoring", self.path)
        except Exception as error:
            self.logger.error("Failed to load resume state %s: %s", self.path, error)
        return {}

    def mark_completed(self, grupo: str, cota: str) -> None:
        if not self.enabled:
            return
        state = self.load_state()
        state['last_processed'] = {'grupo': grupo, 'cota': cota}
        pending = state.get('pending')
        if pending and (
            str(pending.get('grupo', '')) == grupo
            and str(pending.get('cota', '')) == cota
        ):
            state.pop('pending', None)
        state['timestamp'] = datetime.now().isoformat()
        self._write_state(state)

    def mark_pending(self, grupo: str, cota: str) -> None:
        if not self.enabled:
            return
        state = self.load_state()
        state['pending'] = {'grupo': grupo, 'cota': cota}
        state['timestamp'] = datetime.now().isoformat()
        self._write_state(state)

    def clear(self) -> None:
        if not self.enabled or not self.path:
            return
        try:
            if self.path.exists():
                self.path.unlink()
                self.logger.debug("Cleared resume state file %s", self.path)
        except Exception as error:
            self.logger.error("Failed to clear resume state %s: %s", self.path, error)

    def _write_state(self, state: Dict) -> None:
        if not self.enabled or not self.path:
            return
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            tmp_path = self.path.with_suffix(self.path.suffix + '.tmp')
            with open(tmp_path, 'w', encoding='utf-8') as handle:
                json.dump(state, handle, indent=2, ensure_ascii=False)
            tmp_path.replace(self.path)
        except Exception as error:
            self.logger.error("Failed to persist resume state %s: %s", self.path, error)


class FinalWorkingProcessor:
    def __init__(self, config_path: str = "config.yaml"):
        """Initialize the final working processor."""
        self.config_path = Path(config_path).expanduser()
        if not self.config_path.is_absolute():
            self.config_path = Path.cwd() / self.config_path

        self.config = self.load_config(self.config_path)
        self.setup_logging()
        self.setup_directories()
        self.google_drive_uploader: Optional[GoogleDriveUploader] = None
        self.google_sheets_client: Optional[GoogleSheetsClient] = None
        self.google_sheets_logger: Optional[GoogleSheetsClient] = None
        self.google_sheets_log_range: Optional[str] = None
        self.file_link_service: Optional[FileLinkService] = None
        self.notifier: Optional[WebhookNotifier] = None
        self.processed_tracker: Optional[ProcessedRecordTracker] = None
        self.skip_processed_records = False
        self.resume_manager: Optional[ResumeManager] = None
        self.resume_enabled = False
        self.max_login_failures_checkpoint = 5
        self.setup_google_drive()
        self.setup_google_sheets()
        self.setup_google_sheets_logger()
        self.setup_file_server()
        self.setup_notifier()
        self.setup_processed_tracker()
        self.setup_resume_manager()
        
    def load_config(self, config_path: Path) -> Dict:
        """Load configuration from YAML file."""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            return config
        except FileNotFoundError:
            print(f"❌ Configuration file {config_path} not found!")
            sys.exit(1)
        except yaml.YAMLError as e:
            print(f"❌ Error parsing configuration file: {e}")
            sys.exit(1)
    
    def setup_logging(self):
        """Setup logging configuration."""
        log_format = '%(asctime)s - %(levelname)s - %(message)s'
        logging.basicConfig(
            level=logging.INFO,
            format=log_format,
            handlers=[
                logging.FileHandler('final_working_automation.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def setup_directories(self):
        """Create necessary directories."""
        dirs = ['downloads', 'reports', 'screenshots', 'temp']
        for dir_name in dirs:
            Path(dir_name).mkdir(exist_ok=True)

    def setup_google_drive(self):
        drive_config = self.config.get('google_drive', {}) or {}

        if not drive_config.get('enabled', False):
            self.logger.info("Google Drive upload disabled via configuration")
            return

        credentials_path_cfg = drive_config.get('credentials_path')
        drive_id = drive_config.get('drive_id')
        use_year_month = drive_config.get('use_year_month_folders', True)

        if not credentials_path_cfg or not drive_id:
            self.logger.error("Incomplete Google Drive configuration; uploads disabled")
            return

        credentials_path = Path(credentials_path_cfg).expanduser()
        if not credentials_path.is_absolute():
            credentials_path = self.config_path.parent / credentials_path

        delegated_subject = drive_config.get('delegated_subject') or None
        delegated_subject_env = drive_config.get('delegated_subject_env')
        if not delegated_subject and delegated_subject_env:
            env_subject = os.getenv(delegated_subject_env)
            if env_subject:
                delegated_subject = env_subject
                self.logger.info(
                    "Using delegated subject from environment variable %s",
                    delegated_subject_env,
                )
            else:
                self.logger.warning(
                    "Environment variable %s not set; continuing without delegated subject",
                    delegated_subject_env,
                )

        self.google_drive_uploader = GoogleDriveUploader(
            credentials_path=str(credentials_path),
            drive_id=drive_id,
            use_year_month_folders=use_year_month,
            delegated_subject=delegated_subject,
            base_path=self.config_path.parent,
            logger=self.logger,
        )

        if not self.google_drive_uploader.enabled:
            self.logger.warning("Google Drive uploader could not be initialized; continuing without uploads")
        else:
            self.logger.info(
                "Google Drive uploads enabled (folder_id=%s, hierarchy=%s)",
                drive_id,
                "ano/mes" if use_year_month else "flat",
            )

    def setup_google_sheets(self):
        data_source = self.config.get('data_source', {}) or {}
        sheets_config = data_source.get('google_sheets', {}) or {}
        if not sheets_config.get('enabled', False):
            self.logger.info("Google Sheets ingestion disabled via configuration")
            return

        spreadsheet_id = sheets_config.get('spreadsheet_id')
        sheet_range = sheets_config.get('range')
        if not spreadsheet_id or not sheet_range:
            self.logger.error("Google Sheets configuration incomplete; disabling Sheets ingestion")
            return

        credentials_path_cfg = self.config.get('google_drive', {}).get('credentials_path')
        if not credentials_path_cfg:
            self.logger.error("Google Sheets requires service account credentials; none configured")
            return

        credentials_path = Path(credentials_path_cfg).expanduser()
        if not credentials_path.is_absolute():
            credentials_path = self.config_path.parent / credentials_path

        if not credentials_path.exists():
            self.logger.error("Google Sheets credentials file not found: %s", credentials_path)
            return

        self.google_sheets_client = GoogleSheetsClient(
            credentials_path=credentials_path,
            spreadsheet_id=spreadsheet_id,
            logger=self.logger,
        )
        self.logger.info("Google Sheets ingestion enabled (spreadsheet=%s)", spreadsheet_id)

    def setup_google_sheets_logger(self):
        logging_config = self.config.get('google_sheets_logging', {}) or {}
        if not logging_config.get('enabled', False):
            return

        spreadsheet_id = logging_config.get('spreadsheet_id')
        log_range = logging_config.get('range')
        if not spreadsheet_id or not log_range:
            self.logger.error("Google Sheets logging configuration incomplete; disabling logging")
            return

        credentials_path_cfg = self.config.get('google_drive', {}).get('credentials_path')
        if not credentials_path_cfg:
            self.logger.error("Google Sheets logging requires service account credentials")
            return

        credentials_path = Path(credentials_path_cfg).expanduser()
        if not credentials_path.is_absolute():
            credentials_path = self.config_path.parent / credentials_path

        if not credentials_path.exists():
            self.logger.error("Google Sheets logging credentials file not found: %s", credentials_path)
            return

        self.google_sheets_logger = GoogleSheetsClient(
            credentials_path=credentials_path,
            spreadsheet_id=spreadsheet_id,
            logger=self.logger,
            scopes=GoogleSheetsClient.READ_WRITE_SCOPES,
        )
        self.google_sheets_log_range = log_range
        self.logger.info("Google Sheets logging enabled (spreadsheet=%s)", spreadsheet_id)

    def setup_file_server(self):
        file_config = self.config.get('file_server', {}) or {}
        if not file_config.get('enabled', False):
            return

        base_url = file_config.get('base_url')
        secret_key = file_config.get('secret_key')
        downloads_dir_cfg = file_config.get('downloads_dir', 'downloads')
        expiry_minutes = int(file_config.get('expiry_minutes', 30))

        if not base_url or not secret_key:
            self.logger.error("File server configuration incomplete; disabling signed links")
            return

        downloads_dir = Path(downloads_dir_cfg)
        if not downloads_dir.is_absolute():
            downloads_dir = self.config_path.parent / downloads_dir
        downloads_dir.mkdir(parents=True, exist_ok=True)

        try:
            self.file_link_service = FileLinkService(
                downloads_dir=downloads_dir,
                base_url=base_url,
                secret_key=secret_key,
                expiry_minutes=expiry_minutes,
            )
            self.logger.info(
                "Signed file links enabled (base_url=%s, expiry=%s min)",
                base_url,
                expiry_minutes,
            )
        except Exception as error:
            self.logger.error("Failed to initialize file link service: %s", error)
            self.file_link_service = None

    def setup_notifier(self):
        notifications_config = self.config.get('notifications', {}) or {}
        if not notifications_config.get('enabled', False):
            self.logger.info("Notification webhook disabled via configuration")
            return

        webhook_url = notifications_config.get('webhook_url')
        method = notifications_config.get('method', 'POST')
        headers = notifications_config.get('headers')
        template = notifications_config.get('message_template', '')

        if not webhook_url or not template:
            self.logger.error("Notification configuration incomplete; disabling notifier")
            return

        self.notifier = WebhookNotifier(
            webhook_url=webhook_url,
            method=method,
            headers=headers,
            message_template=template,
            logger=self.logger,
        )
        self.logger.info("Notification webhook enabled (%s)", webhook_url)

    def setup_processed_tracker(self) -> None:
        processing_config = self.config.get('processing', {}) or {}
        state_path_cfg = processing_config.get('processed_state_file', 'logs/processed_records.json')
        retention_days = processing_config.get('processed_retention_days')
        skip_processed = processing_config.get('skip_processed_records', False)

        if not state_path_cfg:
            if skip_processed:
                self.logger.warning(
                    "skip_processed_records is enabled but no processed_state_file is configured; disabling skip",
                )
            return

        state_path = Path(state_path_cfg).expanduser()
        if not state_path.is_absolute():
            state_path = self.config_path.parent / state_path

        try:
            self.processed_tracker = ProcessedRecordTracker(
                path=state_path,
                retention_days=retention_days,
                logger=self.logger,
            )
            self.skip_processed_records = bool(skip_processed)
            if self.skip_processed_records:
                self.logger.info(
                    "Processed record tracker enabled; previously completed cotas will be skipped",
                )
            else:
                self.logger.info(
                    "Processed record tracker enabled (tracking only, skipping disabled)",
                )
        except Exception as error:
            self.logger.error("Failed to initialise processed record tracker: %s", error)
            self.processed_tracker = None
            self.skip_processed_records = False

    def setup_resume_manager(self) -> None:
        processing_config = self.config.get('processing', {}) or {}
        resume_enabled = processing_config.get('resume_enabled', True)
        resume_file_cfg = processing_config.get('resume_state_file')
        self.max_login_failures_checkpoint = int(processing_config.get('login_failure_checkpoint', 5))

        if not resume_enabled or not resume_file_cfg:
            return

        resume_path = Path(resume_file_cfg)
        if not resume_path.is_absolute():
            resume_path = self.config_path.parent / resume_path

        self.resume_manager = ResumeManager(resume_path, True, self.logger)
        self.resume_enabled = True
        self.logger.info("Resume manager enabled (state file=%s)", resume_path)

    def parse_submit_function_args(self, onclick_attr: Optional[str]) -> Optional[List[str]]:
        if not onclick_attr:
            return None

        match = re.search(r"submitFunction\((.*)\)", onclick_attr, re.DOTALL)
        if not match:
            return None

        args_str = match.group(1)
        try:
            parsed_args = ast.literal_eval(f"[{args_str}]")
            return ["" if value is None else str(value) for value in parsed_args]
        except (ValueError, SyntaxError):
            return None

    def get_reference_date_from_submit_args(self, submit_args: Optional[List[str]]) -> datetime:
        if submit_args and len(submit_args) > 2:
            due_date_str = submit_args[2]
            try:
                return datetime.strptime(due_date_str, "%d/%m/%Y")
            except (TypeError, ValueError):
                pass
        return datetime.now()

    async def _collect_boleto_form_values(self, page: Page) -> Dict[str, str]:
        result = await page.evaluate(
            """
            () => {
                const form = document.forms.form1;
                if (!form) {
                    return {};
                }
                const fields = {};
                const fieldNames = [
                    'venctoinput',
                    'Data_Limite_Vencimento_Boleto',
                    'FlagAlterarData',
                    'codigo_origem_recurso',
                ];
                for (const name of fieldNames) {
                    const field = form[name];
                    fields[name] = field?.value ?? '';
                }
                return fields;
            }
            """
        )
        return result or {}

    def format_whatsapp_number(self, raw_number: str) -> Optional[str]:
        digits = re.sub(r'\D', '', raw_number or '')
        if not digits:
            return None

        if digits.startswith('55') and len(digits) >= 12:
            return digits

        if len(digits) >= 12:
            return digits

        self.logger.warning("Invalid WhatsApp number detected: %s", raw_number)
        return None

    def sanitize_grupo(self, raw_value: str) -> str:
        if raw_value is None:
            cleaned = ''
        elif isinstance(raw_value, float):
            if math.isnan(raw_value):
                cleaned = ''
            else:
                cleaned = str(int(raw_value)) if raw_value.is_integer() else str(raw_value)
        else:
            cleaned = str(raw_value)
        return re.sub(r'\D', '', cleaned)

    def sanitize_cota(self, raw_value: str) -> str:
        if raw_value is None:
            raw = ''
        elif isinstance(raw_value, float):
            if math.isnan(raw_value):
                raw = ''
            else:
                raw = str(int(raw_value)) if raw_value.is_integer() else str(raw_value)
        else:
            raw = str(raw_value)

        primary_segment = raw.split('-')[0]
        digits = re.sub(r'\D', '', primary_segment)
        if not digits:
            digits = re.sub(r'\D', '', raw)
        return digits

    def load_records(
        self,
        excel_file: str,
        start_from: int,
        max_records: Optional[int],
    ) -> List[Dict]:
        records: List[Dict] = []

        data_source = self.config.get('data_source', {}) or {}
        csv_config = data_source.get('csv', {}) or {}

        if csv_config.get('enabled', False):
            path = csv_config.get('path')
            url = csv_config.get('url')
            encoding = csv_config.get('encoding', 'utf-8')
            delimiter = csv_config.get('delimiter', ',')

            try:
                if path:
                    csv_path = Path(path)
                    if not csv_path.is_absolute():
                        csv_path = self.config_path.parent / csv_path
                    df = pd.read_csv(csv_path, encoding=encoding, sep=delimiter)
                    self.logger.info(f"📊 Loaded {len(df)} records from CSV {csv_path}")
                elif url:
                    response = requests.get(url, timeout=30)
                    response.raise_for_status()
                    df = pd.read_csv(io.StringIO(response.text), encoding=encoding, sep=delimiter)
                    self.logger.info(f"📊 Loaded {len(df)} records from CSV URL")
                else:
                    df = None

                if df is not None:
                    records = df.to_dict('records')
                    for record in records:
                        record['grupo'] = self.sanitize_grupo(record.get('GRUPO') or record.get('grupo'))
                        record['cota'] = self.sanitize_cota(record.get('COTA') or record.get('cota'))
                        record['nome'] = str(record.get('NOME') or record.get('nome') or '').strip()
                        raw_phone = (
                            record.get('WHATS')
                            or record.get('whats')
                            or record.get('telefone')
                            or record.get('TELEFONE')
                            or ''
                        )
                        if isinstance(raw_phone, float):
                            if math.isnan(raw_phone):
                                whats = ''
                            else:
                                whats = str(int(raw_phone))
                        else:
                            whats = str(raw_phone).strip()
                            if whats.endswith('.0'):
                                whats = whats[:-2]
                        record['whats_raw'] = whats
                        record['whats_formatted'] = self.format_whatsapp_number(whats)
                    self.logger.info("Using CSV data source (%d records)", len(records))
            except Exception as error:
                self.logger.error("Failed to load CSV data source: %s", error)

        if not records and self.google_sheets_client:
            sheets_config = data_source.get('google_sheets', {})
            sheet_range = sheets_config.get('range', 'Página1!A:D')
            sheet_rows = self.google_sheets_client.fetch_records(sheet_range)
            for row in sheet_rows:
                record = {
                    'grupo': self.sanitize_grupo(row.get('GRUPO', '')),
                    'cota': self.sanitize_cota(row.get('COTA', '')),
                    'nome': str(row.get('NOME', '')).strip(),
                    'whats_raw': str(
                        row.get('WHATS')
                        or row.get('telefone')
                        or row.get('TELEFONE')
                        or ''
                    ).strip(),
                }
                record['whats_formatted'] = self.format_whatsapp_number(record['whats_raw'])
                records.append(record)

            if not records:
                self.logger.warning("Google Sheets returned no usable records; falling back to Excel")
            else:
                self.logger.info("Using %d records from Google Sheets", len(records))

        if not records:
            df = pd.read_excel(excel_file)
            self.logger.info(f"📊 Loaded {len(df)} records from {excel_file}")
            records = df.to_dict('records')
            for record in records:
                record['grupo'] = self.sanitize_grupo(record.get('grupo') or record.get('GRUPO'))
                record['cota'] = self.sanitize_cota(record.get('cota') or record.get('COTA'))
                raw_phone = (
                    record.get('whats')
                    or record.get('WHATS')
                    or record.get('telefone')
                    or record.get('TELEFONE')
                    or ''
                )
                if isinstance(raw_phone, float):
                    if math.isnan(raw_phone):
                        whats = ''
                    else:
                        whats = str(int(raw_phone))
                else:
                    whats = str(raw_phone).strip()
                    if whats.endswith('.0'):
                        whats = whats[:-2]
                record['whats_raw'] = whats
                record['whats_formatted'] = self.format_whatsapp_number(whats)

        start_index = max(start_from - 1, 0)
        if start_index:
            records = records[start_index:]
        if max_records:
            records = records[:max_records]

        return records

    def handle_post_download(
        self,
        record_info: Dict,
        pdf_path: Path,
        grupo: str,
        cota: str,
        drive_file_id: Optional[str],
    ) -> None:
        phone = record_info.get('whats_formatted')
        nome = record_info.get('nome', 'Cliente')

        file_url: Optional[str] = None
        if self.file_link_service:
            try:
                file_url = self.file_link_service.generate_signed_url(pdf_path)
            except Exception as error:
                self.logger.error("Failed to generate signed URL for %s: %s", pdf_path, error)

        notification_success: Optional[bool] = None
        if self.notifier:
            if phone:
                notification_success = self.notifier.send_notification(
                    phone_number=phone,
                    nome=nome,
                    grupo=grupo,
                    cota=cota,
                    pdf_path=Path(pdf_path),
                    file_url=file_url,
                    drive_file_id=drive_file_id,
                )

                if notification_success:
                    self.logger.info("📨 Notification sent to %s for %s/%s", phone, grupo, cota)
                else:
                    self.logger.warning("⚠️ Notification failed for %s/%s", grupo, cota)
            else:
                self.logger.warning(
                    "Skipping notification for %s/%s due to missing or invalid WhatsApp number",
                    grupo,
                    cota,
                )
            if file_url is None:
                self.logger.error("Signed link generation failed; notification may not reach the client")
        elif file_url:
            self.logger.debug("Notifier disabled; generated signed URL %s", file_url)
        else:
            self.logger.warning("Notifier disabled and no signed URL available for %s", pdf_path.name)

        if self.google_sheets_logger:
            self.log_processing_result(
                grupo=grupo,
                cota=cota,
                nome=nome,
                phone=phone or record_info.get('whats_raw'),
                pdf_path=pdf_path,
                drive_file_id=drive_file_id,
                file_url=file_url,
                notification_success=notification_success,
            )

    def log_processing_result(
        self,
        grupo: str,
        cota: str,
        nome: str,
        phone: Optional[str],
        pdf_path: Path,
        drive_file_id: Optional[str],
        file_url: Optional[str],
        notification_success: Optional[bool],
    ) -> None:
        if not self.google_sheets_logger or not self.google_sheets_log_range:
            return

        timestamp = datetime.now().isoformat()
        drive_status = "OK" if drive_file_id else "SKIPPED"
        notification_status = (
            "OK" if notification_success else "FAIL" if notification_success is False else "SKIPPED"
        )
        values = [
            timestamp,
            grupo,
            cota,
            nome,
            phone or '',
            str(pdf_path),
            drive_file_id or '',
            drive_status,
            notification_status,
            file_url or '',
        ]

        self.google_sheets_logger.append_row(self.google_sheets_log_range, values)
    
    def generate_filename(self, nome: str, grupo: str, cota: str, cpf_cnpj: str, index: int = 0) -> str:
        """Generate safe filename for boleto PDF."""
        nome_clean = re.sub(r'[^\w\s-]', '', nome.strip())[:20] if nome else 'CLIENTE'
        nome_clean = re.sub(r'\s+', '-', nome_clean)
        cpf_cnpj_clean = re.sub(r'[^\d]', '', cpf_cnpj) if cpf_cnpj else 'UNKNOWN'
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{nome_clean}-{grupo}-{cota}-{cpf_cnpj_clean}-{timestamp}-{index}.pdf"
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        return filename

    async def open_boleto_page_directly(self, page: Page, onclick_attr: str) -> Optional[Page]:
        """Parses onclick to construct and open the boleto URL directly."""
        try:
            context = page.context
            self.logger.info("🔧 Parsing onclick to open boleto page directly.")
            match = re.search(r"submitFunction\((.*)\)", onclick_attr)
            if not match:
                self.logger.error("❌ Could not parse submitFunction parameters.")
                return None

            # Robustly parse parameters, handling commas inside quotes
            params_str = match.group(1)
            # This regex splits by comma, but ignores commas inside single quotes
            params = re.findall(r"'([^']*)'", params_str)
            
            if len(params) < 14:
                self.logger.error(f"❌ Incorrect parameter count after parsing. Expected 14+, got {len(params)}.")
                return None

            # Mapping based on typical form submissions for Slip.asp
            # Replicate the full form submission, including hidden fields
            form_data = {
                'codigo_agente': params[0],
                'numero_aviso': params[1],
                'vencto': params[2],
                'descricao': params[3],
                'codigo_grupo': params[4],
                'codigo_cota': params[5],
                'codigo_movimento': params[6],
                'valor_total': params[7].replace(',', '.'), # CRITICAL FIX: Ensure decimal is a period
                'desc_pagamento': params[8],
                'msg_dbt_apenas_parc_antes_venc': params[10],
                'sn_emite_boleto_pix': params[13],
                # Include other hidden fields from the form
                'venctoinput': '', # This is set to null by the JS if empty
                'Data_Limite_Vencimento_Boleto': '', # Assuming this is not critical or is set server-side
                'FlagAlterarData': 'N',
                'codigo_origem_recurso': '0'
            }

            slip_url = self.config['site']['base_url'] + 'Slip/Slip.asp'
            self.logger.info(f"🚀 Submitting POST request to: {slip_url}")

            # Use a new page to perform the POST request by building a temporary form
            temp_page = await context.new_page()
            
            # Listen for the new page (boleto) to be created by the form submission
            async with context.expect_page() as new_page_info:
                await temp_page.evaluate("""(args) => {
                    const form = document.createElement('form');
                    form.method = 'POST';
                    form.action = args.url;
                    form.target = '_blank'; // Ensures submission opens in a new tab

                    for (const key in args.data) {
                        const input = document.createElement('input');
                        input.type = 'hidden';
                        input.name = key;
                        input.value = args.data[key];
                        form.appendChild(input);
                    }

                    document.body.appendChild(form);
                    form.submit();
                }""", {'url': slip_url, 'data': form_data})
            
            boleto_page = await new_page_info.value
            await temp_page.close() # Clean up the temporary page

            await boleto_page.wait_for_load_state('domcontentloaded', timeout=30000)
            
            content = await boleto_page.content()
            if len(content) > 1000 and 'ADODB.Command' not in content:
                self.logger.info(f"✅ Successfully loaded boleto page via form submission with {len(content)} characters.")
                return boleto_page
            else:
                self.logger.error(f"❌ Form submission resulted in an error page or empty content.")
                await boleto_page.close()
                return None

        except Exception as e:
            self.logger.error(f"❌ Error opening boleto page directly: {e}")
            return None
    
    async def wait_for_pdf_generation(self, pdf_path: str, timeout: float = 60.0, min_size: int = 20000) -> bool:
        """Wait for PDF generation to complete by monitoring file size."""
        try:
            self.logger.info(f"⏰ WAITING FOR PDF GENERATION: {pdf_path}")
            
            start_time = time.time()
            last_size = 0
            stable_count = 0
            
            while (time.time() - start_time) < timeout:
                if Path(pdf_path).exists():
                    current_size = Path(pdf_path).stat().st_size
                    self.logger.info(f"⏰ PDF size: {current_size} bytes (was {last_size})")
                    
                    if current_size >= min_size:
                        if current_size == last_size:
                            stable_count += 1
                            if stable_count >= 3:
                                self.logger.info(f"✅ PDF GENERATION COMPLETE: {current_size} bytes")
                                return True
                        else:
                            stable_count = 0
                    
                    last_size = current_size
                
                await asyncio.sleep(2)
            
            if Path(pdf_path).exists():
                final_size = Path(pdf_path).stat().st_size
                self.logger.warning(f"⚠️ PDF timeout, final size: {final_size} bytes")
                return final_size >= min_size
            else:
                self.logger.error(f"❌ PDF file never created: {pdf_path}")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ Error waiting for PDF: {e}")
            return False
    
    async def final_working_pgto_parc_click(self, page: Page, link, index: int) -> Optional[Page]:
        """FINAL WORKING METHOD: Open boleto page directly from onclick attribute."""
        try:
            self.logger.info(f"🚀 FINAL WORKING METHOD for boleto {index}")
            
            onclick = await link.get_attribute('onclick')
            if not onclick:
                self.logger.error("❌ No onclick attribute found")
                return None

            self.logger.info(f"🔍 onclick: {onclick}")

            boleto_page = await self.open_boleto_page_directly(page, onclick)

            if boleto_page:
                self.logger.info(f"✅ FINAL SUCCESS: Boleto page loaded with URL: {boleto_page.url}")
                return boleto_page
            else:
                self.logger.error(f"❌ FINAL FAILURE: Could not load boleto content for boleto {index}")
                return None
            
        except Exception as e:
            self.logger.error(f"❌ Final working method failed: {e}")
            return None
    
    async def login(self, page: Page) -> bool:
        """Login to the system."""
        try:
            self.logger.info("Starting login process...")
            
            await page.goto(self.config['site']['base_url'], timeout=30000)
            await page.wait_for_load_state('domcontentloaded')
            await asyncio.sleep(2)
            
            iframe_element = await page.wait_for_selector('iframe', timeout=10000)
            iframe = await iframe_element.content_frame()
            
            if not iframe:
                self.logger.error("Could not access iframe content!")
                return False
            
            await iframe.fill("input[name='j_username']", self.config['login']['username'])
            await iframe.fill("input[name='j_password']", self.config['login']['password'])
            await asyncio.sleep(1)
            await iframe.click("input[name='btnLogin']")
            
            await page.wait_for_load_state('domcontentloaded', timeout=15000)
            await asyncio.sleep(2)
            
            self.logger.info("✅ Login successful")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Login failed: {e}")
            return False
    
    async def search_record(self, page: Page, grupo: str, cota: str) -> Tuple[bool, Dict]:
        """Search for a specific grupo/cota record."""
        try:
            self.logger.info(f"Searching for Grupo: {grupo}, Cota: {cota}")
            
            search_url = self.config['site']['search_url']
            await page.goto(search_url, timeout=30000)
            await page.wait_for_load_state('domcontentloaded')
            await asyncio.sleep(2)
            
            # Handle frames
            frames = page.frames
            search_frame = page
            for frame in frames:
                if 'searchCota' in frame.url or 'Attendance' in frame.url:
                    search_frame = frame
                    break
            
            await search_frame.fill("input[name='Grupo']", grupo)
            await search_frame.fill("input[name='Cota']", cota)
            await asyncio.sleep(1)
            await search_frame.click("input[name='Button']")
            
            await page.wait_for_load_state('domcontentloaded', timeout=15000)
            await asyncio.sleep(3)
            
            # Extract CPF/CNPJ and status
            current_url = page.url
            cpf_cnpj = None
            if 'cgc_cpf_cliente=' in current_url:
                cpf_cnpj = current_url.split('cgc_cpf_cliente=')[1].split('&')[0]
            
            # Detect contemplado status
            page_content = await page.content()
            page_upper = page_content.upper()
            contemplado_status = "UNKNOWN"

            contemplado_keywords = self.config['contemplado']['keywords']['contemplado']
            nao_contemplado_keywords = self.config['contemplado']['keywords']['nao_contemplado']

            # First check for explicit "não contemplado" style phrases so we do not
            # misclassify due to the substring "CONTEMPLADO".
            for keyword in nao_contemplado_keywords:
                if keyword.upper() in page_upper:
                    contemplado_status = "NÃO CONTEMPLADO"
                    break

            if contemplado_status == "UNKNOWN":
                for keyword in contemplado_keywords:
                    if keyword.upper() in page_upper:
                        contemplado_status = "CONTEMPLADO"
                        break
            
            result = {
                'cpf_cnpj': cpf_cnpj,
                'contemplado_status': contemplado_status,
                'page_url': current_url
            }
            
            self.logger.info(f"✅ Search successful - CPF/CNPJ: {cpf_cnpj}, Status: {contemplado_status}")
            return True, result
            
        except Exception as e:
            self.logger.error(f"❌ Search failed for {grupo}/{cota}: {e}")
            return False, {'error': str(e)}
    
    async def download_boletos_final_working(self, page: Page, grupo: str, cota: str, record_info: Dict, timing_config: Dict) -> List[str]:
        """FINAL WORKING VERSION: Download boletos with proper submitFunction execution."""
        downloaded_files = []
        
        try:
            self.logger.info(f"🚀 FINAL WORKING BOLETO DOWNLOAD for {grupo}/{cota}")
            
            # Find and click 2ª Via Boleto
            segunda_via_links = await page.query_selector_all("a[title*='2ª Via Boleto'], a[href*='emissSlip.asp']")
            if not segunda_via_links:
                self.logger.warning("No 2ª Via Boleto links found")
                return downloaded_files
            
            self.logger.info("Clicking 2ª Via Boleto link")
            await segunda_via_links[0].click()
            await asyncio.sleep(timing_config.get('segunda_via_delay', 3))
            
            # Populate boleto table by entering due date and clicking Salvar
            self.logger.info("Populating boleto table...")
            
            # Wait for the boleto generation form to load
            try:
                await page.wait_for_selector("input[name='venctoinput']:not([type='hidden']), input[type='text'][size='10']", timeout=10000)
            except:
                self.logger.warning("Could not find visible date input field, trying alternative approach")
            
            # Fill in due date (30 days from now)
            due_date = (datetime.now() + timedelta(days=30)).strftime("%d/%m/%Y")
            
            # Try different selectors for the visible due date input
            date_input = None
            selectors_to_try = [
                "input[name='venctoinput']:not([type='hidden'])",
                "input[type='text'][size='10']",
                "input[type='text'][maxlength='10']",
                "input[type='text'][placeholder*='data']",
                "input[type='text'][name*='venc']"
            ]
            
            for selector in selectors_to_try:
                date_input = await page.query_selector(selector)
                if date_input:
                    # Check if it's visible
                    is_visible = await date_input.is_visible()
                    if is_visible:
                        self.logger.info(f"Found visible date input with selector: {selector}")
                        break
                    else:
                        date_input = None
                        
            if date_input:
                await date_input.fill('')  # Clear the field
                await date_input.fill(due_date)
                self.logger.info(f"Filled due date: {due_date}")
            else:
                self.logger.warning("Could not find visible due date input field")
                # Save debug HTML to see the form structure
                debug_html = await page.content()
                debug_path = f"downloads/debug_form_{grupo}_{cota}.html"
                with open(debug_path, 'w', encoding='utf-8') as f:
                    f.write(debug_html)
                self.logger.info(f"Saved form debug HTML: {debug_path}")
            
            # Click Salvar button to populate the table
            salvar_selectors = [
                "input[value='Salvar']",
                "input[type='submit'][value*='Salvar']", 
                "button:has-text('Salvar')",
                "input[type='button'][value*='Salvar']"
            ]
            
            salvar_button = None
            for selector in salvar_selectors:
                salvar_button = await page.query_selector(selector)
                if salvar_button:
                    is_visible = await salvar_button.is_visible()
                    if is_visible:
                        self.logger.info(f"Found Salvar button with selector: {selector}")
                        break
                    else:
                        salvar_button = None
                        
            if salvar_button:
                await salvar_button.click()
                self.logger.info("Clicked Salvar button")
                await asyncio.sleep(3)  # Wait for table to populate
            else:
                self.logger.warning("Could not find Salvar button")
            
            # Find PGTO PARC links after table population
            pgto_parc_links = await page.query_selector_all("a[href*='javascript:'][onclick*='submitFunction'], a:has-text('PGTO PARC')")
            if not pgto_parc_links:
                self.logger.warning("No PGTO PARC links found after table population")
                # Save debug HTML
                debug_html = await page.content()
                debug_path = f"downloads/debug_no_pgto_parc_{grupo}_{cota}.html"
                with open(debug_path, 'w', encoding='utf-8') as f:
                    f.write(debug_html)
                self.logger.info(f"Saved debug HTML: {debug_path}")
                return downloaded_files
            
            self.logger.info(f"Found {len(pgto_parc_links)} PGTO PARC links")
            
            # Determine how many boletos to download
            contemplado_status = record_info.get('contemplado_status', 'UNKNOWN')
            if contemplado_status == "CONTEMPLADO":
                links_to_process = pgto_parc_links[:1]
                self.logger.info("CONTEMPLADO - downloading most recent boleto only")
            else:
                links_to_process = pgto_parc_links[:1]
                self.logger.info("NÃO CONTEMPLADO - downloading most recent boleto only")
            
            # Process each PGTO PARC link with direct POST method
            for i, link in enumerate(links_to_process):
                try:
                    self.logger.info(f"🚀 PROCESSING BOLETO {i+1}/{len(links_to_process)} - DIRECT POST METHOD")
                    
                    onclick_attr = await link.get_attribute('onclick')
                    if not onclick_attr:
                        self.logger.error("No onClick attribute found for PGTO PARC link")
                        continue

                    submit_args = self.parse_submit_function_args(onclick_attr)
                    if not submit_args:
                        self.logger.error("Unable to parse submitFunction arguments for boleto %s", i + 1)
                        continue

                    self.logger.info(f"📋 onClick: {onclick_attr}")

                    # Extract onClick parameters and make direct POST request
                    pdf_data = await self.extract_and_fetch_boleto_direct(
                        page,
                        link,
                        i + 1,
                        onclick_attr=onclick_attr,
                        submit_args=submit_args,
                    )

                    if not pdf_data:
                        self.logger.error(f"❌ FAILED TO GET PDF DATA for boleto {i+1}")
                        continue

                    self.logger.info(f"✅ PDF DATA RECEIVED: {len(pdf_data)} bytes")
                    
                    # Generate filename
                    nome = record_info.get('nome', 'CLIENTE')
                    cpf_cnpj = record_info.get('cpf_cnpj', 'UNKNOWN')
                    filename = self.generate_filename(nome, grupo, cota, cpf_cnpj, i)
                    pdf_path = f'downloads/{filename}'
                    
                    # Save PDF data to file
                    try:
                        with open(pdf_path, 'wb') as f:
                            f.write(pdf_data)
                        
                        # Verify file was created and has content
                        drive_file_id: Optional[str] = None
                        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 10000:
                            downloaded_files.append(pdf_path)
                            file_size = os.path.getsize(pdf_path)
                            self.logger.info(f"✅ BOLETO {i+1} DOWNLOADED: {filename} ({file_size} bytes)")

                            reference_date = self.get_reference_date_from_submit_args(submit_args)

                            if self.google_drive_uploader and self.google_drive_uploader.enabled:
                                drive_file_id = self.google_drive_uploader.upload_pdf(
                                    local_path=pdf_path,
                                    file_name=filename,
                                    reference_date=reference_date,
                                )
                                if drive_file_id:
                                    self.logger.info(
                                        "📁 BOLETO %s uploaded to Google Drive (file_id=%s)",
                                        filename,
                                        drive_file_id,
                                    )
                                    record_info.setdefault('drive_file_ids', []).append(drive_file_id)
                                else:
                                    self.logger.warning(
                                        "⚠️ Google Drive upload failed for %s",
                                        filename,
                                    )

                            self.handle_post_download(record_info, Path(pdf_path), grupo, cota, drive_file_id)
                        else:
                            self.logger.error(f"❌ PDF file too small or missing: {filename}")
                            
                    except Exception as save_error:
                        self.logger.error(f"❌ Failed to save PDF {i+1}: {save_error}")
                        
                except Exception as e:
                    self.logger.error(f"❌ Error processing boleto {i+1}: {e}")
                    
            return downloaded_files
            
        except Exception as e:
            self.logger.error(f"❌ Download process failed: {e}")
            return downloaded_files
    
    async def extract_and_fetch_boleto_direct(
        self,
        page: Page,
        link,
        boleto_num: int,
        onclick_attr: Optional[str] = None,
        submit_args: Optional[List[str]] = None,
    ) -> Optional[bytes]:
        """Extract onClick parameters and make direct POST request to get PDF blob."""
        try:
            self.logger.info(f"🔍 Extracting onClick parameters for boleto {boleto_num}")

            if onclick_attr is None:
                onclick_attr = await link.get_attribute('onclick')

            if submit_args is None:
                submit_args = self.parse_submit_function_args(onclick_attr)

            if not onclick_attr or not submit_args:
                self.logger.error("Unable to process boleto %s due to missing submitFunction data", boleto_num)
                return None

            form_values = await self._collect_boleto_form_values(page)
            action_url = await page.evaluate(
                "() => new URL('../Slip/Slip.asp', window.location.href).toString()"
            )

            try:
                (
                    codigo_agente,
                    numero_aviso,
                    vencto,
                    descricao,
                    codigo_grupo,
                    codigo_cota,
                    codigo_movimento,
                    valor_total,
                    desc_pagamento,
                    debito_conta,
                    msg_boleto,
                    emite_mensagem_ident_cob,
                    vSN_Emite_Boleto,
                    vSN_Emite_Boleto_Pix,
                ) = submit_args[:14]
            except ValueError:
                self.logger.error("Unexpected submitFunction argument format for boleto %s", boleto_num)
                return None

            payload = {
                'numero_aviso': numero_aviso,
                'vencto': vencto,
                'venctoinput': form_values.get('venctoinput', ''),
                'valor_total': valor_total,
                'descricao': descricao,
                'codigo_grupo': codigo_grupo,
                'codigo_cota': codigo_cota,
                'codigo_movimento': codigo_movimento,
                'codigo_agente': codigo_agente,
                'desc_pagamento': desc_pagamento,
                'msg_dbt_apenas_parc_antes_venc': msg_boleto,
                'sn_emite_boleto_pix': vSN_Emite_Boleto_Pix,
                'Data_Limite_Vencimento_Boleto': form_values.get('Data_Limite_Vencimento_Boleto', ''),
                'FlagAlterarData': form_values.get('FlagAlterarData', 'N'),
                'codigo_origem_recurso': form_values.get('codigo_origem_recurso', '0'),
            }

            response = await page.context.request.post(
                action_url,
                form=payload,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
            )

            if not response.ok:
                self.logger.error(
                    "Slip POST request failed for boleto %s: HTTP %s",
                    boleto_num,
                    response.status,
                )
                return None

            pdf_bytes = await response.body()
            self.logger.info(f"✅ Got PDF data: {len(pdf_bytes)} bytes")
            return pdf_bytes

        except Exception as e:
            self.logger.error(f"❌ Error in extract_and_fetch_boleto_direct: {e}")
            return None
    
    async def process_record(self, browser: Browser, record: Dict, timing_config: Dict) -> Dict:
        """Process a single record with final working method."""
        context = await browser.new_context(
            viewport={'width': 1280, 'height': 720},
            accept_downloads=True
        )
        
        page = await context.new_page()
        
        grupo = str(record.get('grupo', '')).strip()
        cota = str(record.get('cota', '')).strip()
        nome = record.get('nome', 'UNKNOWN')
        whats_raw = record.get('whats_raw') or record.get('whats') or ''
        whats_formatted = record.get('whats_formatted')

        result = {
            'grupo': grupo,
            'cota': cota,
            'nome': nome,
            'whats_raw': whats_raw,
            'whats_formatted': whats_formatted,
            'status': 'failed',
            'downloaded_files': [],
            'timestamp': datetime.now().isoformat()
        }
        
        try:
            self.logger.info(f"Processing record: {grupo}/{cota} - {nome}")
            
            # Login
            if not await self.login(page):
                result['status'] = 'login_failed'
                return result
            
            # Search
            search_success, search_result = await self.search_record(page, grupo, cota)
            if not search_success:
                result['status'] = 'search_failed'
                result.update(search_result)
                return result
            
            result.update(search_result)
            result['nome'] = nome
            result['whats_raw'] = whats_raw
            result['whats_formatted'] = whats_formatted

            result['whats_raw'] = whats_raw
            result['whats_formatted'] = whats_formatted
            search_result['nome'] = nome
            search_result['whats_raw'] = whats_raw
            search_result['whats_formatted'] = whats_formatted

            # Download with final working method
            downloaded_files = await self.download_boletos_final_working(page, grupo, cota, search_result, timing_config)
            result['downloaded_files'] = downloaded_files
            result['downloaded_count'] = len(downloaded_files)
            if search_result.get('drive_file_ids'):
                result['drive_file_ids'] = list(search_result['drive_file_ids'])
            
            if downloaded_files:
                result['status'] = 'success'
                self.logger.info(f"✅ SUCCESS: {grupo}/{cota} - {len(downloaded_files)} files")
            else:
                result['status'] = 'no_downloads'
                self.logger.warning(f"⚠️ NO DOWNLOADS: {grupo}/{cota}")
            
        except Exception as e:
            result['status'] = 'error'
            result['error'] = str(e)
            self.logger.error(f"❌ Error processing {grupo}/{cota}: {e}")
        
        finally:
            await context.close()
        
        return result
    
    def setup_resume_manager(self) -> None:
        processing_config = self.config.get('processing', {}) or {}
        resume_enabled = processing_config.get('resume_enabled', True)
        resume_file_cfg = processing_config.get('resume_state_file')
        self.max_login_failures_checkpoint = int(processing_config.get('login_failure_checkpoint', 5))

        if not resume_enabled or not resume_file_cfg:
            return

        resume_path = Path(resume_file_cfg)
        if not resume_path.is_absolute():
            resume_path = self.config_path.parent / resume_path

        self.resume_manager = ResumeManager(resume_path, True, self.logger)
        self.resume_enabled = True
        self.logger.info("Resume manager enabled (state file=%s)", resume_path)


    @staticmethod
    def find_record_index(records: List[Dict], key: Dict[str, str]) -> Optional[int]:
        if not key:
            return None
        target_grupo = str(key.get('grupo', '')).strip()
        target_cota = str(key.get('cota', '')).strip()
        if not target_grupo or not target_cota:
            return None
        for idx, record in enumerate(records):
            if (
                str(record.get('grupo', '')).strip() == target_grupo
                and str(record.get('cota', '')).strip() == target_cota
            ):
                return idx
        return None

    async def run_automation(self, excel_file: str, start_from: int = 0, max_records: int = None, batch_size: int = 100, timing_config: Dict = None, ignore_resume: bool = False):
        """Run the final working automation."""
        if timing_config is None:
            timing_config = {
                'popup_delay': 5.0,
                'content_delay': 5.0,
                'pre_pdf_delay': 6.0,
                'post_pdf_delay': 3.0,
                'segunda_via_delay': 3.0,
                'pdf_wait_timeout': 60.0,
                'min_pdf_size': 20000
            }
        
        try:
            records = self.load_records(excel_file, start_from, max_records)
            if not records:
                self.logger.error("No records available to process. Aborting run.")
                return

            skipped_count = 0
            if self.processed_tracker and self.skip_processed_records:
                filtered_records = []
                for record in records:
                    grupo = str(record.get('grupo', '')).strip()
                    cota = str(record.get('cota', '')).strip()
                    if grupo and cota and self.processed_tracker.is_processed(grupo, cota):
                        skipped_count += 1
                        self.logger.info(
                            "⏭️ Skipping already processed record %s/%s",
                            grupo,
                            cota,
                        )
                        continue
                    filtered_records.append(record)
                if skipped_count:
                    self.logger.info(
                        "⏭️ Skipped %d records that were already processed in previous runs",
                        skipped_count,
                    )
                records = filtered_records

            if self.resume_manager and self.resume_enabled and not ignore_resume and start_from == 0:
                resume_state = self.resume_manager.load_state()
                pending = resume_state.get('pending') if resume_state else None
                if pending:
                    idx = self.find_record_index(records, pending)
                    if idx is not None:
                        records = records[idx:]
                        self.logger.info(
                            "Resuming from pending record %s/%s (position %s of %s)",
                            pending.get('grupo'),
                            pending.get('cota'),
                            idx + 1,
                            len(records),
                        )
                    else:
                        self.logger.warning("Pending record %s/%s not found; clearing resume state", pending.get('grupo'), pending.get('cota'))
                        self.resume_manager.clear()
                elif resume_state and resume_state.get('last_processed'):
                    last_key = resume_state['last_processed']
                    idx = self.find_record_index(records, last_key)
                    if idx is not None:
                        records = records[idx + 1:]
                        self.logger.info(
                            "Resuming after record %s/%s (skipping %s entries)",
                            last_key.get('grupo'),
                            last_key.get('cota'),
                            idx + 1,
                        )
                    else:
                        self.logger.info("Resume checkpoint already processed; starting with remaining records")

            if not records:
                self.logger.info("No new records to process after applying processed-record filter.")
                return

            if start_from > 0:
                self.logger.info(f"📍 Starting from record {start_from}")
            if max_records:
                self.logger.info(f"📊 Limited to {max_records} records")

            self.logger.info(f"🎯 Processing {len(records)} records")
            self.logger.info(f"🚀 FINAL WORKING VERSION: submitFunction in main page context")
            
            # Launch browser
            async with async_playwright() as p:
                browser = await p.chromium.launch(
                    headless=True,
                    slow_mo=1000,
                    args=[
                        '--no-sandbox',
                        '--disable-setuid-sandbox',
                        '--disable-dev-shm-usage',
                        '--disable-gpu',
                        '--disable-web-security',
                        '--disable-background-timer-throttling',
                        '--disable-renderer-backgrounding',
                        '--disable-popup-blocking',
                        '--print-to-pdf-no-header',
                        '--run-all-compositor-stages-before-draw'
                    ]
                )
                
                all_results = []
                total_downloads = 0
                consecutive_login_failures = 0
                
                for i in range(0, len(records), batch_size):
                    batch = records[i:i + batch_size]
                    batch_num = (i // batch_size) + 1
                    
                    self.logger.info(f"🚀 Batch {batch_num} ({len(batch)} records)")
                    
                    for j, record in enumerate(batch, 1):
                        self.logger.info(f"Record {j}/{len(batch)} in batch {batch_num}")
                        result = await self.process_record(browser, record, timing_config)
                        all_results.append(result)
                        total_downloads += result.get('downloaded_count', 0)

                        grupo_key = str(result.get('grupo', '')).strip()
                        cota_key = str(result.get('cota', '')).strip()

                        status = result.get('status')
                        if self.processed_tracker and status in ('success', 'no_downloads'):
                            drive_ids = result.get('drive_file_ids')
                            metadata = {
                                'timestamp': result.get('timestamp'),
                                'downloaded_files': result.get('downloaded_files', []) if status == 'success' else [],
                                'cpf_cnpj': result.get('cpf_cnpj'),
                                'status': status,
                            }
                            if drive_ids:
                                metadata['drive_file_ids'] = drive_ids
                            self.processed_tracker.mark_processed(
                                grupo_key,
                                cota_key,
                                metadata,
                            )

                        if self.resume_manager and self.resume_enabled and grupo_key and cota_key:
                            if status == 'login_failed':
                                self.resume_manager.mark_pending(grupo_key, cota_key)
                            else:
                                self.resume_manager.mark_completed(grupo_key, cota_key)

                        if status == 'login_failed':
                            consecutive_login_failures += 1
                            if consecutive_login_failures >= self.max_login_failures_checkpoint:
                                self.logger.error(
                                    'Exceeded %s consecutive login failures; checkpoint reached. Stopping run for later retry.',
                                    self.max_login_failures_checkpoint,
                                )
                                return
                        else:
                            consecutive_login_failures = 0

                        # Save intermediate results
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        with open(f'reports/final_working_results_{timestamp}.json', 'w', encoding='utf-8') as f:
                            json.dump(all_results, f, indent=2, ensure_ascii=False)

                        await asyncio.sleep(5)
                    
                    # Between batches pause
                    if i + batch_size < len(records):
                        self.logger.info(f"⏸️ Pausing 20s between batches...")
                        await asyncio.sleep(20)
                
                await browser.close()

            if self.resume_manager and self.resume_enabled:
                self.resume_manager.clear()
            
            # Final summary
            successful = len([r for r in all_results if r['status'] == 'success'])
            failed = len([r for r in all_results if r['status'] not in ['success', 'no_downloads']])
            no_downloads = len([r for r in all_results if r['status'] == 'no_downloads'])
            
            self.logger.info("🎉 FINAL WORKING AUTOMATION COMPLETED!")
            self.logger.info(f"📊 Summary: {successful} successful, {failed} failed, {no_downloads} no downloads")
            self.logger.info(f"📁 Total files: {total_downloads}")
            
            # Save final report
            final_report = {
                'summary': {
                    'total_records': len(all_results),
                    'successful': successful,
                    'failed': failed,
                    'no_downloads': no_downloads,
                    'total_downloads': total_downloads,
                    'success_rate': round((successful/len(all_results)*100), 2) if all_results else 0,
                    'timing_config': timing_config
                },
                'results': all_results,
                'timestamp': datetime.now().isoformat()
            }
            
            final_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            with open(f'reports/final_working_report_{final_timestamp}.json', 'w', encoding='utf-8') as f:
                json.dump(final_report, f, indent=2, ensure_ascii=False)
            
            print(f"\n🚀 FINAL WORKING RESULTS:")
            print(f"   Total Records: {len(all_results)}")
            print(f"   Successful: {successful}")
            print(f"   Failed: {failed}")
            print(f"   No Downloads: {no_downloads}")
            print(f"   Total Files: {total_downloads}")
            print(f"   Success Rate: {final_report['summary']['success_rate']}%")
            
        except Exception as e:
            self.logger.error(f"❌ Final working automation failed: {e}")
            raise


def main():
    """Main entry point for final working automation."""
    parser = argparse.ArgumentParser(
        description='Final Working Boleto Automation - Execute submitFunction in main page context'
    )
    
    parser.add_argument('excel_file', help='Excel file containing boleto data')
    parser.add_argument('--start-from', type=int, default=0, help='Start from record number')
    parser.add_argument('--max-records', type=int, default=None, help='Max records to process')
    parser.add_argument('--batch-size', type=int, default=100, help='Batch size')
    parser.add_argument('--config', default='config.yaml', help='Config file path')
    parser.add_argument('--ignore-resume', action='store_true', help='Ignore resume checkpoints and start fresh')
    
    # Timing options
    parser.add_argument('--popup-delay', type=float, default=5.0, help='Popup delay')
    parser.add_argument('--content-delay', type=float, default=5.0, help='Content delay')
    parser.add_argument('--pre-pdf-delay', type=float, default=6.0, help='Pre-PDF delay')
    parser.add_argument('--pdf-wait-timeout', type=float, default=60.0, help='PDF timeout')
    parser.add_argument('--min-pdf-size', type=int, default=20000, help='Min PDF size')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.excel_file):
        print(f"❌ Excel file not found: {args.excel_file}")
        sys.exit(1)
    
    if not os.path.exists(args.config):
        print(f"❌ Config file not found: {args.config}")
        sys.exit(1)
    
    timing_config = {
        'popup_delay': args.popup_delay,
        'content_delay': args.content_delay,
        'pre_pdf_delay': args.pre_pdf_delay,
        'post_pdf_delay': 3.0,
        'segunda_via_delay': 3.0,
        'pdf_wait_timeout': args.pdf_wait_timeout,
        'min_pdf_size': args.min_pdf_size
    }
    
    try:
        processor = FinalWorkingProcessor(args.config)
        asyncio.run(processor.run_automation(
            excel_file=args.excel_file,
            start_from=args.start_from,
            max_records=args.max_records,
            batch_size=args.batch_size,
            timing_config=timing_config,
            ignore_resume=args.ignore_resume
        ))
        
    except KeyboardInterrupt:
        print("\n⏹️ Automation interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Automation failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
