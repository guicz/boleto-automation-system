#!/usr/bin/env python3
"""
Enhanced Production Boleto Processor
Integrates the successful direct POST approach with production-ready automation
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


class EnhancedProductionProcessor:
    def __init__(self, config_path: str = "config.yaml"):
        """Initialize the enhanced production processor."""
        self.config = self.load_config(config_path)
        self.setup_logging()
        self.setup_directories()
        
    def load_config(self, config_path: str) -> Dict:
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
        """Enhanced boleto download using direct POST approach."""
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
                    
                    # Extract onClick parameters and make direct POST request
                    pdf_data = await self.extract_and_fetch_boleto_direct(page, link, i+1)
                    
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
                    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 10000:
                        downloaded_files.append(pdf_path)
                        file_size = os.path.getsize(pdf_path)
                        self.logger.info(f"✅ BOLETO {i+1} DOWNLOADED: {filename} ({file_size} bytes)")
                    else:
                        self.logger.error(f"❌ PDF file too small or missing: {filename}")
                        
                except Exception as save_error:
                    self.logger.error(f"❌ Failed to save PDF {i+1}: {save_error}")
                    
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
                    try {{
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
                        
                    }} catch (error) {{
                        console.error("Error in JavaScript:", error);
                        return null;
                    }}
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
        """Process a single record with enhanced method."""
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
        
        # Load Excel data
        try:
            df = pd.read_excel(excel_file)
            self.logger.info(f"📊 Loaded {len(df)} records from {excel_file}")
        except Exception as e:
            self.logger.error(f"❌ Error loading Excel file: {e}")
            return
        
        # Apply filters
        if start_from > 1:
            df = df.iloc[start_from-1:]
            self.logger.info(f"📍 Starting from record {start_from}")
        
        if max_records:
            df = df.head(max_records)
            self.logger.info(f"📊 Limited to {max_records} records")
        
        records = df.to_dict('records')
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
