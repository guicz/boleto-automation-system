#!/usr/bin/env python3
"""
Debug SubmitFunction Test
Test the submitFunction execution in main page context to diagnose popup loading issues
"""

import asyncio
import logging
import sys
from datetime import datetime
from pathlib import Path
import yaml
from playwright.async_api import async_playwright

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('debug_submitfunction.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

async def debug_direct_navigation():
    """Debug the direct navigation boleto fetching method."""
    import re
    from urllib.parse import urlencode

    # Load config
    try:
        with open('config.yaml', 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
    except FileNotFoundError:
        logger.error("❌ config.yaml not found!")
        return False
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,  # Must run headless in this environment
            slow_mo=1000,
        )
        context = await browser.new_context()
        page = await context.new_page()
        
        try:
            # Step 1: Login
            logger.info("🔄 Step 1: Login")
            await page.goto(config['site']['base_url'], timeout=30000)
            iframe = page.frame_locator('iframe').locator("input[name='j_username']")
            await iframe.fill(config['login']['username'])
            await page.frame_locator('iframe').locator("input[name='j_password']").fill(config['login']['password'])
            await page.frame_locator('iframe').locator("input[name='btnLogin']").click()
            await page.wait_for_load_state('domcontentloaded', timeout=15000)
            logger.info("✅ Login successful")

            # Step 2: Search for a test record
            logger.info("🔄 Step 2: Search for test record")
            await page.goto(config['site']['search_url'], timeout=30000)
            # Use a known test case
            test_grupo = "684"
            test_cota = "644"
            search_frame = page.frame(url=re.compile("searchCota.asp")) or page
            await search_frame.fill("input[name='Grupo']", test_grupo)
            await search_frame.fill("input[name='Cota']", test_cota)
            await search_frame.click("input[name='Button']")
            await page.wait_for_load_state('domcontentloaded', timeout=15000)
            logger.info(f"✅ Search completed for {test_grupo}/{test_cota}")

            # Step 3: Click 2ª Via Boleto
            logger.info("🔄 Step 3: Click 2ª Via Boleto")
            await page.locator("a[title*='2ª Via Boleto'], a[href*='emissSlip.asp']").first.click()
            await asyncio.sleep(3) # Wait for page to update
            logger.info("✅ 2ª Via Boleto clicked")

            # Step 4: Extract onclick from the first PGTO PARC link
            logger.info("🔄 Step 4: Extract onclick attribute")
            first_link = page.locator("a[href*='javascript:'][onclick*='submitFunction']").first
            onclick_attr = await first_link.get_attribute('onclick')
            if not onclick_attr:
                logger.error("❌ No onclick attribute found on the first PGTO PARC link.")
                return False
            logger.info(f"✅ Found onclick attribute: {onclick_attr}")

            # Step 5: Parse parameters and construct URL
            logger.info("🔄 Step 5: Parse parameters and construct URL")
            match = re.search(r"submitFunction\((.*)\)", onclick_attr)
            if not match:
                logger.error("❌ Could not parse submitFunction parameters.")
                return False
            
            params_str = match.group(1)
            params = re.findall(r"'([^']*)'", params_str)
            logger.info(f"✅ Parsed {len(params)} parameters.")

            if len(params) < 14:
                logger.error(f"❌ Incorrect parameter count after parsing. Expected 14+, got {len(params)}.")
                return False

            form_data = {
                'codigo_agente': params[0],
                'numero_aviso': params[1],
                'vencto': params[2],
                'descricao': params[3],
                'codigo_grupo': params[4],
                'codigo_cota': params[5],
                'codigo_movimento': params[6],
                'valor_total': params[7].replace(',', '.'), # CRITICAL FIX
                'desc_pagamento': params[8],
                'msg_dbt_apenas_parc_antes_venc': params[10],
                'sn_emite_boleto_pix': params[13],
                'venctoinput': '',
                'Data_Limite_Vencimento_Boleto': '',
                'FlagAlterarData': 'N',
                'codigo_origem_recurso': '0'
            }
            
            slip_url = config['site']['base_url'] + 'Slip/Slip.asp'
            logger.info(f"🚀 Submitting POST request to: {slip_url}")

            # Step 6: Send authenticated POST request and load content
            logger.info("🔄 Step 6: Send authenticated POST request and verify content")
            
            cookies = await page.context.cookies()
            cookie_header = "; ".join([f"{c['name']}={c['value']}" for c in cookies])

            api_request_context = page.request
            response = await api_request_context.post(
                slip_url,
                form=form_data,
                headers={
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Cookie': cookie_header,
                    'Referer': page.url
                }
            )

            if not response.ok:
                logger.error(f"❌ POST request failed with status {response.status}: {response.status_text}")
                return False

            response_body = await response.body()
            boleto_html = response_body.decode('iso-8859-1')

            if len(boleto_html) < 1000 or 'ADODB.Command' in boleto_html:
                logger.error(f"❌ POST response content indicates an error ({len(boleto_html)} chars).")
                logger.debug(f"Response content: {boleto_html}")
                return False

            boleto_page = await page.context.new_page()
            await boleto_page.set_content(boleto_html, wait_until='domcontentloaded')
            
            content = await boleto_page.content()
            content_length = len(content)
            logger.info(f"📄 Boleto page content length: {content_length} characters.")

            if content_length > 1000:
                logger.info("🎉 SUCCESS! Direct navigation loaded boleto content.")
                screenshot_path = f'screenshots/debug_direct_nav_success_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png'
                await boleto_page.screenshot(path=screenshot_path, full_page=True)
                logger.info(f"📸 Success screenshot saved to {screenshot_path}")
                return True
            else:
                logger.error("❌ Direct navigation resulted in a page with insufficient content.")
                screenshot_path = f'screenshots/debug_direct_nav_failure_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png'
                await boleto_page.screenshot(path=screenshot_path, full_page=True)
                logger.info(f"📸 Failure screenshot saved to {screenshot_path}")
                return False

        except Exception as e:
            logger.error(f"❌ An error occurred during the debug test: {e}", exc_info=True)
            return False
        
        finally:
            logger.info("🔍 Keeping browser open for 20 seconds for manual inspection...")
            await asyncio.sleep(20)
            await browser.close()

async def main():
    """Main function."""
    logger.info("🚀 Starting direct navigation debug test...")
    
    # Ensure required directories exist
    Path('screenshots').mkdir(exist_ok=True)
    Path('temp').mkdir(exist_ok=True)

    success = await debug_direct_navigation()
    
    if success:
        logger.info("🎉 DEBUG TEST SUCCESSFUL - Direct navigation method is working!")
    else:
        logger.error("❌ DEBUG TEST FAILED - Direct navigation method failed to load boleto.")
    
    return success

if __name__ == "__main__":
    try:
        result = asyncio.run(main())
        sys.exit(0 if result else 1)
    except KeyboardInterrupt:
        print("\n⏹️ Debug test interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Debug test failed: {e}")
        sys.exit(1)
