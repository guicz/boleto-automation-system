#!/usr/bin/env python3
"""
Enhanced Production Boleto Processor v2.0.0
Integrates the successful direct POST approach with production-ready automation

Features:
- Direct POST method for reliable PDF extraction from HS Consórcios system
- Iframe-aware login handling
- Frame-based search functionality  
- Customer name-based file naming from Excel data
- Batch processing with configurable delays
- Comprehensive logging and JSON report generation
- Fixed duplicate boleto downloads issue

Version History:
- v1.0.0: Initial implementation with basic functionality
- v2.0.0: Fixed duplicate downloads, enhanced selectors, production-ready

Author: Boleto Automation Team
Created: 2025-09-05
Last Updated: 2025-09-05
"""

__version__ = "2.0.0"
__author__ = "Boleto Automation Team"
__email__ = "automation@company.com"
__status__ = "Production"

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


class EnhancedProductionProcessor:
    """
    Enhanced Production Boleto Processor - Main automation class.
    
    This class handles the complete workflow for downloading boleto PDFs:
    1. Login to HS Consórcios system via iframe
    2. Search for grupo/cota records in frames
    3. Navigate to boleto generation pages
    4. Populate boleto tables with due dates
    5. Extract parameters from PGTO PARC links
    6. Make direct POST requests to retrieve PDF blobs
    7. Save PDFs with customer-based naming convention
    """
    
    def __init__(self, config_path: str = "config.yaml"):
        """Initialize the enhanced production processor.
        
        Args:
            config_path (str): Path to YAML configuration file
        """
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
        self.setup_google_drive()
        self.setup_google_sheets()
        self.setup_google_sheets_logger()
        self.setup_file_server()
        self.setup_notifier()

    def load_config(self, config_path: Path) -> Dict:
        """Load configuration from YAML file."""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            print(f"❌ Error loading config: {e}")
            sys.exit(1)
    
    def setup_logging(self):
        """Setup logging configuration."""
        log_level = getattr(logging, self.config.get('logging', {}).get('level', 'INFO').upper())
        
        # Create logs directory
        os.makedirs('logs', exist_ok=True)
        
        # Setup logging
        logging.basicConfig(
            level=log_level,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(f'logs/enhanced_automation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
    def setup_directories(self):
        """Create necessary directories."""
        directories = ['downloads', 'screenshots', 'reports', 'temp', 'logs']
        for directory in directories:
            os.makedirs(directory, exist_ok=True)

    def setup_google_drive(self):
        """Initialize Google Drive uploader if configuration is provided."""
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
            self.logger.error(
                "Google Sheets credentials file not found: %s",
                credentials_path,
            )
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
            self.logger.error(
                "Google Sheets logging credentials file not found: %s",
                credentials_path,
            )
            return

        self.google_sheets_logger = GoogleSheetsClient(
            credentials_path=credentials_path,
            spreadsheet_id=spreadsheet_id,
            logger=self.logger,
            scopes=GoogleSheetsClient.READ_WRITE_SCOPES,
        )
        self.google_sheets_log_range = logging_config.get('range')
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
        digits = re.sub(r'\D', '', raw_value or '')
        return digits

    def sanitize_cota(self, raw_value: str) -> str:
        value = (raw_value or '').split('-')[0]
        digits = re.sub(r'\D', '', value)
        if not digits:
            digits = re.sub(r'\D', '', raw_value or '')
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
            if start_from > 0:
                df = df.iloc[start_from:]
            if max_records:
                df = df.head(max_records)
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
    
    async def login(self, page: Page) -> bool:
        """Login to the HS Consórcios system using iframe-based authentication.
        
        The login form is embedded in an iframe, so we need to:
        1. Navigate to the base URL
        2. Wait for iframe to load
        3. Access iframe content
        4. Fill username and password fields
        5. Submit login form
        
        Args:
            page (Page): Playwright page instance
            
        Returns:
            bool: True if login successful, False otherwise
        """
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
    
    async def search_grupo_cota(self, page: Page, grupo: str, cota: str) -> Tuple[bool, Dict]:
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
            contemplado_status = "UNKNOWN"
            
            # Simple contemplado detection
            if "CONTEMPLADO" in page_content:
                if "NÃO CONTEMPLADO" in page_content:
                    contemplado_status = "NÃO CONTEMPLADO"
                else:
                    contemplado_status = "CONTEMPLADO"
            
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
    
    async def extract_record_info(self, page: Page) -> Dict:
        """Extract record information from search results."""
        try:
            # Extract CPF/CNPJ
            cpf_cnpj = "UNKNOWN"
            cpf_elements = await page.query_selector_all("td:has-text('CPF'), td:has-text('CNPJ')")
            if cpf_elements:
                for element in cpf_elements:
                    text = await element.text_content()
                    if text and ('CPF' in text or 'CNPJ' in text):
                        # Extract numbers from the text
                        numbers = re.findall(r'\d+', text)
                        if numbers:
                            cpf_cnpj = ''.join(numbers)
                            break
            
            # Extract contemplado status
            contemplado_status = "UNKNOWN"
            page_content = await page.content()
            if "CONTEMPLADO" in page_content:
                if "NÃO CONTEMPLADO" in page_content:
                    contemplado_status = "NÃO CONTEMPLADO"
                else:
                    contemplado_status = "CONTEMPLADO"
            
            return {
                'cpf_cnpj': cpf_cnpj,
                'contemplado_status': contemplado_status
            }
            
        except Exception as e:
            self.logger.error(f"❌ Error extracting record info: {e}")
            return {'cpf_cnpj': 'UNKNOWN', 'contemplado_status': 'UNKNOWN'}
    
    async def download_boletos_enhanced(self, page: Page, grupo: str, cota: str, record_info: Dict, timing_config: Dict) -> List[str]:
        """Enhanced boleto download using direct POST approach.
        
        This method implements the core boleto download logic:
        1. Click '2ª Via Boleto' link to navigate to generation page
        2. Populate boleto table by filling due date and clicking 'Salvar'
        3. Find PGTO PARC links using specific CSS selector
        4. Extract onClick parameters from each link
        5. Make direct POST requests to Slip.asp endpoint
        6. Save PDF blobs with customer-based filenames
        
        Args:
            page (Page): Playwright page instance
            grupo (str): Grupo number
            cota (str): Cota number
            record_info (Dict): Record information including name, CPF/CNPJ, status
            timing_config (Dict): Timing configuration for delays
            
        Returns:
            List[str]: List of downloaded file paths
        """
        downloaded_files = []
        
        try:
            self.logger.info(f"🚀 ENHANCED BOLETO DOWNLOAD for {grupo}/{cota}")
            
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
            # Fixed: Using single specific selector to prevent duplicate selections
            # Previous comma-separated selector was matching same links twice
            pgto_parc_links = await page.query_selector_all("a[href*='javascript:'][onclick*='submitFunction']:has-text('PGTO PARC')")
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
                links_to_process = pgto_parc_links
                self.logger.info(f"NÃO CONTEMPLADO - downloading all {len(links_to_process)} boletos")
            
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
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    
                    # Clean nome for filename
                    nome_clean = re.sub(r'[^\w\s-]', '', nome).strip()
                    nome_clean = re.sub(r'[-\s]+', '-', nome_clean).upper()
                    
                    filename = f"{nome_clean}-{grupo}-{cota}-{cpf_cnpj}-{timestamp}-{i}.pdf"
                    pdf_path = f"downloads/{filename}"
                    
                    # Save PDF
                    with open(pdf_path, 'wb') as f:
                        f.write(pdf_data)

                    # Verify file was created and has content
                    drive_file_id: Optional[str] = None
                    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 10000:
                        downloaded_files.append(pdf_path)
                        file_size = os.path.getsize(pdf_path)
                        self.logger.info(f"✅ BOLETO {i+1} DOWNLOADED: {filename} ({file_size} bytes)")

                        reference_date = self.get_reference_date_from_submit_args(submit_args)

                        # Upload to Google Drive if enabled
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
        """Extract onClick parameters and make direct POST request to get PDF blob.

        This method performs the critical PDF extraction:
        1. Get onClick attribute from PGTO PARC link
        2. Parse JavaScript submitFunction parameters
        3. Build form data for POST request
        4. Execute direct POST to Slip.asp with session cookies
        5. Return PDF blob as bytes
        
        Args:
            page (Page): Playwright page instance
            link: HTML link element with onClick attribute
            boleto_num (int): Boleto number for logging
            
        Returns:
            bytes: PDF content as bytes, or None if failed
        """
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
        """Process a single record with enhanced method."""
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
                result['error'] = 'Login failed'
                return result
            
            # Search
            search_success, record_info = await self.search_grupo_cota(page, grupo, cota)
            if not search_success:
                result['error'] = 'Search failed'
                return result
            
            # Add record info to result and preserve original nome
            result.update(record_info)
            result['nome'] = nome  # Preserve original nome
            
            # Add nome to record_info for filename generation
            record_info['nome'] = nome
            record_info['whats_raw'] = whats_raw
            record_info['whats_formatted'] = whats_formatted

            # Download boletos
            downloaded_files = await self.download_boletos_enhanced(page, grupo, cota, record_info, timing_config)
            
            if downloaded_files:
                result['status'] = 'success'
                result['downloaded_files'] = downloaded_files
                self.logger.info(f"✅ SUCCESS: {grupo}/{cota} - {len(downloaded_files)} files")
            else:
                result['status'] = 'no_downloads'
                self.logger.warning(f"⚠️ NO DOWNLOADS: {grupo}/{cota}")
            
        except Exception as e:
            self.logger.error(f"❌ Error processing {grupo}/{cota}: {e}")
            result['error'] = str(e)
            
        finally:
            await context.close()
            await asyncio.sleep(2)  # Brief pause between records
            
        return result
    
    async def run_automation(self, excel_file: str, start_from: int = 1, max_records: Optional[int] = None, 
                           batch_size: int = 1, timing_config: Dict = None) -> None:
        """Run the enhanced automation process."""
        
        records = self.load_records(excel_file, start_from, max_records)
        if not records:
            self.logger.error("No records available to process. Aborting run.")
            return

        if start_from > 1:
            self.logger.info(f"📍 Starting from record {start_from}")
        if max_records:
            self.logger.info(f"📊 Limited to {max_records} records")

        self.logger.info(f"🎯 Processing {len(records)} records")
        self.logger.info("🚀 ENHANCED PRODUCTION VERSION: Direct POST with table population")
        
        # Default timing config
        if timing_config is None:
            timing_config = {
                'segunda_via_delay': 3.0,
                'popup_delay': 2.0,
                'content_delay': 1.0,
                'pre_pdf_delay': 1.0,
                'post_pdf_delay': 2.0
            }
        
        # Initialize results tracking
        results = []
        successful = 0
        failed = 0
        no_downloads = 0
        total_files = 0
        
        async with async_playwright() as p:
            browser = await p.chromium.launch(
                headless=self.config.get('browser', {}).get('headless', True),
                slow_mo=self.config.get('browser', {}).get('slow_mo', 500)
            )
            
            try:
                # Process records in batches
                for batch_start in range(0, len(records), batch_size):
                    batch_end = min(batch_start + batch_size, len(records))
                    batch_records = records[batch_start:batch_end]
                    batch_num = (batch_start // batch_size) + 1
                    
                    self.logger.info(f"🚀 Batch {batch_num} ({len(batch_records)} records)")
                    
                    # Process each record in the batch
                    for i, record in enumerate(batch_records):
                        record_num = batch_start + i + 1
                        self.logger.info(f"Record {record_num}/{len(records)} in batch {batch_num}")
                        
                        result = await self.process_record(browser, record, timing_config)
                        results.append(result)
                        
                        # Update counters
                        if result['status'] == 'success':
                            successful += 1
                            total_files += len(result['downloaded_files'])
                        elif result['status'] == 'failed':
                            failed += 1
                        else:  # no_downloads
                            no_downloads += 1
                        
                        # Brief pause between records
                        await asyncio.sleep(1)
                
            finally:
                await browser.close()
        
        # Generate summary
        success_rate = (successful / len(records)) * 100 if records else 0
        
        self.logger.info("🎉 ENHANCED PRODUCTION AUTOMATION COMPLETED!")
        self.logger.info(f"📊 Summary: {successful} successful, {failed} failed, {no_downloads} no downloads")
        self.logger.info(f"📁 Total files: {total_files}")
        
        # Save results report
        report_path = f"reports/enhanced_automation_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        os.makedirs('reports', exist_ok=True)
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump({
                'summary': {
                    'total_records': len(records),
                    'successful': successful,
                    'failed': failed,
                    'no_downloads': no_downloads,
                    'total_files': total_files,
                    'success_rate': success_rate
                },
                'results': results,
                'timestamp': datetime.now().isoformat()
            }, f, indent=2, ensure_ascii=False)
        
        self.logger.info(f"📄 Report saved: {report_path}")
        
        # Print final summary
        print(f"\n🚀 ENHANCED PRODUCTION RESULTS:")
        print(f"   Total Records: {len(records)}")
        print(f"   Successful: {successful}")
        print(f"   Failed: {failed}")
        print(f"   No Downloads: {no_downloads}")
        print(f"   Total Files: {total_files}")
        print(f"   Success Rate: {success_rate:.1f}%")


def main():
    """Main function with argument parsing."""
    parser = argparse.ArgumentParser(description='Enhanced Production Boleto Processor')
    parser.add_argument('excel_file', help='Path to Excel file with grupo/cota data')
    parser.add_argument('--start-from', type=int, default=1, help='Start from record number (1-based)')
    parser.add_argument('--max-records', type=int, help='Maximum number of records to process')
    parser.add_argument('--batch-size', type=int, default=1, help='Number of records per batch')
    parser.add_argument('--config', default='config.yaml', help='Configuration file path')
    parser.add_argument('--segunda-via-delay', type=float, default=3.0, help='Delay after clicking 2ª Via Boleto')
    parser.add_argument('--popup-delay', type=float, default=2.0, help='Delay for popup handling')
    parser.add_argument('--content-delay', type=float, default=1.0, help='Delay for content loading')
    
    args = parser.parse_args()
    
    # Validate Excel file
    if not os.path.exists(args.excel_file):
        print(f"❌ Excel file not found: {args.excel_file}")
        sys.exit(1)
    
    # Timing configuration
    timing_config = {
        'segunda_via_delay': args.segunda_via_delay,
        'popup_delay': args.popup_delay,
        'content_delay': args.content_delay,
        'pre_pdf_delay': 1.0,
        'post_pdf_delay': 2.0
    }
    
    try:
        processor = EnhancedProductionProcessor(args.config)
        asyncio.run(processor.run_automation(
            excel_file=args.excel_file,
            start_from=args.start_from,
            max_records=args.max_records,
            batch_size=args.batch_size,
            timing_config=timing_config
        ))
        
    except KeyboardInterrupt:
        print("\n⏹️ Automation interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Automation failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
