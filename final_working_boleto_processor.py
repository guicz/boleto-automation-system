#!/usr/bin/env python3
"""
Final Working Boleto Processor
SOLUTION: Parse the onclick attribute to construct and submit a POST request directly to Slip.asp, 
replicating the website's form submission for maximum reliability.
"""

import asyncio
import argparse
import json
import logging
import os
import sys
import re
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import yaml
from playwright.async_api import async_playwright, Browser, Page, BrowserContext


class FinalWorkingProcessor:
    def __init__(self, config_path: str = "config.yaml"):
        """Initialize the final working processor."""
        self.config = self.load_config(config_path)
        self.setup_logging()
        self.setup_directories()
        
    def load_config(self, config_path: str) -> Dict:
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
            contemplado_status = "UNKNOWN"
            
            contemplado_keywords = self.config['contemplado']['keywords']['contemplado']
            nao_contemplado_keywords = self.config['contemplado']['keywords']['nao_contemplado']
            
            for keyword in contemplado_keywords:
                if keyword in page_content.upper():
                    contemplado_status = "CONTEMPLADO"
                    break
            
            if contemplado_status == "UNKNOWN":
                for keyword in nao_contemplado_keywords:
                    if keyword in page_content.upper():
                        contemplado_status = "NÃO CONTEMPLADO"
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
                links_to_process = pgto_parc_links
                self.logger.info(f"NÃO CONTEMPLADO - downloading all {len(links_to_process)} boletos")
            
            # Process each PGTO PARC link with direct POST method
            for i, link in enumerate(links_to_process):
                try:
                    self.logger.info(f"🚀 PROCESSING BOLETO {i+1}/{len(links_to_process)} - DIRECT POST METHOD")
                    
                    # Extract onClick parameters and make direct POST request
                    pdf_data = await self.extract_and_fetch_boleto_direct(page, link, i+1)
                    
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
                        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 10000:
                            downloaded_files.append(pdf_path)
                            file_size = os.path.getsize(pdf_path)
                            self.logger.info(f"✅ BOLETO {i+1} DOWNLOADED: {filename} ({file_size} bytes)")
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
    
    async def extract_and_fetch_boleto_direct(self, page: Page, link, boleto_num: int) -> bytes:
        """Extract onClick parameters and make direct POST request to get PDF blob."""
        try:
            self.logger.info(f"🔍 Extracting onClick parameters for boleto {boleto_num}")
            
            # Get the onClick attribute
            onclick_attr = await link.get_attribute('onclick')
            if not onclick_attr:
                self.logger.error("No onClick attribute found")
                return None
                
            self.logger.info(f"📋 onClick: {onclick_attr}")
            
            # Execute JavaScript to extract parameters and make POST request
            pdf_data = await page.evaluate(f"""
                async () => {{
                    // Extract parameters from onClick attribute
                    const onClickStr = `{onclick_attr}`;
                    const match = onClickStr.match(/submitFunction\\(([\\s\\S]*)\\)/);
                    if (!match) {{
                        console.error("Could not parse onClick parameters");
                        return null;
                    }}
                    
                    // Parse arguments
                    const argsStr = match[1].replace(/'/g, '"');
                    const args = JSON.parse('[' + argsStr + ']');
                    const [
                        ca, na, v, d, cg, cc, cm, vt,
                        desc_pagamento, debito_conta, msg_boleto,
                        emite_mensagem_ident_cob, vSN_Emite_Boleto, vSN_Emite_Boleto_Pix
                    ] = args;
                    
                    // Get form fields
                    const form = document.forms.form1;
                    const venctoinput = form?.venctoinput?.value || "";
                    const Data_Limite_Vencimento_Boleto = form?.Data_Limite_Vencimento_Boleto?.value || "";
                    const FlagAlterarData = form?.FlagAlterarData?.value || "N";
                    const codigo_origem_recurso = form?.codigo_origem_recurso?.value || "0";
                    
                    // Build form data
                    const formData = new URLSearchParams({{
                        numero_aviso: na,
                        vencto: v,
                        venctoinput: venctoinput,
                        valor_total: vt,
                        descricao: d,
                        codigo_grupo: cg,
                        codigo_cota: cc,
                        codigo_movimento: cm,
                        codigo_agente: ca,
                        desc_pagamento: desc_pagamento,
                        msg_dbt_apenas_parc_antes_venc: msg_boleto,
                        sn_emite_boleto_pix: vSN_Emite_Boleto_Pix,
                        Data_Limite_Vencimento_Boleto: Data_Limite_Vencimento_Boleto,
                        FlagAlterarData: FlagAlterarData,
                        codigo_origem_recurso: codigo_origem_recurso
                    }});
                    
                    // Make POST request
                    const actionUrl = new URL("../Slip/Slip.asp", location.href).toString();
                    console.log("Making POST request to:", actionUrl);
                    
                    const response = await fetch(actionUrl, {{
                        method: "POST",
                        credentials: "include",
                        headers: {{ "Content-Type": "application/x-www-form-urlencoded" }},
                        body: formData.toString()
                    }});
                    
                    if (!response.ok) {{
                        throw new Error("HTTP error: " + response.status);
                    }}
                    
                    // Get response as array buffer
                    const arrayBuffer = await response.arrayBuffer();
                    return Array.from(new Uint8Array(arrayBuffer));
                }}
            """)
            
            if not pdf_data:
                self.logger.error("Failed to get PDF data from JavaScript")
                return None
                
            # Convert array back to bytes
            pdf_bytes = bytes(pdf_data)
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
        
        grupo = str(record['grupo'])
        cota = str(record['cota'])
        nome = record.get('nome', 'UNKNOWN')
        
        result = {
            'grupo': grupo,
            'cota': cota,
            'nome': nome,
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
            
            # Download with final working method
            downloaded_files = await self.download_boletos_final_working(page, grupo, cota, result, timing_config)
            result['downloaded_files'] = downloaded_files
            result['downloaded_count'] = len(downloaded_files)
            
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
    
    async def run_automation(self, excel_file: str, start_from: int = 0, max_records: int = None, batch_size: int = 5, timing_config: Dict = None):
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
            # Load data
            df = pd.read_excel(excel_file)
            self.logger.info(f"📊 Loaded {len(df)} records from {excel_file}")
            
            # Apply filters
            if start_from > 0:
                df = df.iloc[start_from:]
                self.logger.info(f"📍 Starting from record {start_from}")
            
            if max_records:
                df = df.head(max_records)
                self.logger.info(f"📊 Limited to {max_records} records")
            
            records = df.to_dict('records')
            
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
                
                for i in range(0, len(records), batch_size):
                    batch = records[i:i + batch_size]
                    batch_num = (i // batch_size) + 1
                    
                    self.logger.info(f"🚀 Batch {batch_num} ({len(batch)} records)")
                    
                    for j, record in enumerate(batch, 1):
                        self.logger.info(f"Record {j}/{len(batch)} in batch {batch_num}")
                        result = await self.process_record(browser, record, timing_config)
                        all_results.append(result)
                        total_downloads += result.get('downloaded_count', 0)
                        
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
    parser.add_argument('--batch-size', type=int, default=5, help='Batch size')
    parser.add_argument('--config', default='config.yaml', help='Config file path')
    
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
