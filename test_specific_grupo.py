#!/usr/bin/env python3
"""
Test script for specific Grupo 001148 / Cota 0479
Goes to the page after clicking 2ª Via Boleto and runs our direct POST approach
"""

import asyncio
import logging
import os
import yaml
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

async def test_specific_grupo():
    """Test with Grupo 001148 / Cota 0479"""
    
    # Load config
    with open('config.yaml', 'r') as f:
        config = yaml.safe_load(f)
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            viewport={'width': 1280, 'height': 720},
            accept_downloads=True
        )
        page = await context.new_page()
        
        try:
            # Login
            logger.info("Starting login process...")
            await page.goto(config['site']['search_url'])
            await page.fill(config['selectors']['login']['username'], config['login']['username'])
            await page.fill(config['selectors']['login']['password'], config['login']['password'])
            await page.click(config['selectors']['login']['submit'])
            await asyncio.sleep(3)
            logger.info("✅ Login successful")
            
            # Search for specific grupo/cota
            grupo = "001148"
            cota = "0479"
            logger.info(f"Searching for Grupo: {grupo}, Cota: {cota}")
            
            await page.fill(config['selectors']['search']['grupo'], grupo)
            await page.fill(config['selectors']['search']['cota'], cota)
            await page.click(config['selectors']['search']['submit'])
            await asyncio.sleep(5)
            logger.info("✅ Search completed")
            
            # Click 2ª Via Boleto
            logger.info("Looking for 2ª Via Boleto link...")
            segunda_via_links = await page.query_selector_all("a[title*='2ª Via Boleto'], a[href*='emissSlip.asp']")
            if not segunda_via_links:
                logger.error("No 2ª Via Boleto links found")
                return
            
            logger.info("Clicking 2ª Via Boleto link")
            await segunda_via_links[0].click()
            await asyncio.sleep(5)
            
            # Now we're on the boleto generation page - populate the table
            logger.info("🎯 NOW ON BOLETO GENERATION PAGE - RUNNING OUR NEW CODE")
            
            # Fill in due date and click Salvar to populate table
            logger.info("Populating boleto table...")
            
            # Try to find visible date input
            date_input = None
            selectors_to_try = [
                "input[name='venctoinput']:not([type='hidden'])",
                "input[type='text'][size='10']",
                "input[type='text'][maxlength='10']",
                "input[type='text'][name*='venc']"
            ]
            
            for selector in selectors_to_try:
                date_input = await page.query_selector(selector)
                if date_input:
                    is_visible = await date_input.is_visible()
                    if is_visible:
                        logger.info(f"Found visible date input with selector: {selector}")
                        break
                    else:
                        date_input = None
            
            if date_input:
                due_date = (datetime.now() + timedelta(days=30)).strftime("%d/%m/%Y")
                await date_input.fill('')  # Clear
                await date_input.fill(due_date)
                logger.info(f"Filled due date: {due_date}")
            else:
                logger.warning("Could not find visible due date input")
                # Save debug HTML
                debug_html = await page.content()
                with open('debug_form.html', 'w', encoding='utf-8') as f:
                    f.write(debug_html)
                logger.info("Saved debug HTML: debug_form.html")
            
            # Click Salvar button
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
                        logger.info(f"Found Salvar button with selector: {selector}")
                        break
                    else:
                        salvar_button = None
            
            if salvar_button:
                await salvar_button.click()
                logger.info("Clicked Salvar button")
                await asyncio.sleep(5)  # Wait for table to populate
            else:
                logger.warning("Could not find Salvar button")
            
            # Look for PGTO PARC links
            pgto_parc_links = await page.query_selector_all("a[href*='javascript:'][onclick*='submitFunction'], a:has-text('PGTO PARC')")
            logger.info(f"Found {len(pgto_parc_links)} PGTO PARC links")
            
            if not pgto_parc_links:
                logger.warning("No PGTO PARC links found - saving debug HTML")
                debug_html = await page.content()
                with open('debug_no_pgto_parc.html', 'w', encoding='utf-8') as f:
                    f.write(debug_html)
                logger.info("Saved debug HTML: debug_no_pgto_parc.html")
                return
            
            # Test our direct POST approach with the first PGTO PARC link
            link = pgto_parc_links[0]
            logger.info("🚀 TESTING DIRECT POST APPROACH")
            
            # Get the onClick attribute
            onclick_attr = await link.get_attribute('onclick')
            logger.info(f"📋 onClick: {onclick_attr}")
            
            if not onclick_attr:
                logger.error("No onClick attribute found")
                return
            
            # Execute our JavaScript to extract parameters and make POST request
            pdf_data = await page.evaluate(f"""
                async () => {{
                    try {{
                        // Extract parameters from onClick attribute
                        const onClickStr = `{onclick_attr}`;
                        console.log("onClick string:", onClickStr);
                        
                        const match = onClickStr.match(/submitFunction\\(([\\s\\S]*)\\)/);
                        if (!match) {{
                            console.error("Could not parse onClick parameters");
                            return null;
                        }}
                        
                        // Parse arguments
                        const argsStr = match[1].replace(/'/g, '"');
                        console.log("Args string:", argsStr);
                        
                        const args = JSON.parse('[' + argsStr + ']');
                        console.log("Parsed args:", args);
                        
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
                        
                        console.log("Form data:", formData.toString());
                        
                        // Make POST request
                        const actionUrl = new URL("../Slip/Slip.asp", location.href).toString();
                        console.log("Making POST request to:", actionUrl);
                        
                        const response = await fetch(actionUrl, {{
                            method: "POST",
                            credentials: "include",
                            headers: {{ "Content-Type": "application/x-www-form-urlencoded" }},
                            body: formData.toString()
                        }});
                        
                        console.log("Response status:", response.status);
                        console.log("Response headers:", Object.fromEntries(response.headers.entries()));
                        
                        if (!response.ok) {{
                            throw new Error("HTTP error: " + response.status);
                        }}
                        
                        // Get response as array buffer
                        const arrayBuffer = await response.arrayBuffer();
                        console.log("Response size:", arrayBuffer.byteLength, "bytes");
                        
                        return Array.from(new Uint8Array(arrayBuffer));
                        
                    }} catch (error) {{
                        console.error("Error in JavaScript:", error);
                        return null;
                    }}
                }}
            """)
            
            if pdf_data:
                pdf_bytes = bytes(pdf_data)
                logger.info(f"✅ Got PDF data: {len(pdf_bytes)} bytes")
                
                # Save PDF
                filename = f"boleto_{grupo}_{cota}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                os.makedirs('downloads', exist_ok=True)
                pdf_path = f"downloads/{filename}"
                
                with open(pdf_path, 'wb') as f:
                    f.write(pdf_bytes)
                
                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
                    logger.info(f"🎉 SUCCESS! PDF saved: {filename} ({os.path.getsize(pdf_path)} bytes)")
                else:
                    logger.error(f"❌ PDF file too small or missing: {filename}")
            else:
                logger.error("❌ Failed to get PDF data")
            
            # Keep browser open for inspection
            logger.info("🔍 Keeping browser open for inspection...")
            await asyncio.sleep(30)
            
        except Exception as e:
            logger.error(f"❌ Error: {e}")
            
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(test_specific_grupo())
