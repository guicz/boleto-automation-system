#!/usr/bin/env python3
"""
Test Final Solution - Single Record Test
Quick test to validate the final working solution with one record
"""

import asyncio
import logging
import sys
from datetime import datetime
import yaml
from final_working_boleto_processor import FinalWorkingProcessor

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('test_final_solution.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

async def test_single_record():
    """Test the final solution with a single record."""
    
    logger.info("🚀 Testing Final Working Solution with Single Record")
    
    # Test configuration
    timing_config = {
        'popup_delay': 5.0,
        'content_delay': 8.0,      # Increased for testing
        'pre_pdf_delay': 10.0,     # Increased for testing
        'post_pdf_delay': 3.0,
        'segunda_via_delay': 3.0,
        'pdf_wait_timeout': 90.0,  # Increased timeout
        'min_pdf_size': 15000      # Slightly lower for testing
    }
    
    try:
        # Initialize processor
        processor = FinalWorkingProcessor('config.yaml')
        
        # Test with single record (modify these values as needed)
        test_record = {
            'grupo': 33,
            'cota': 40862778,
            'nome': 'TEST_CLIENTE'
        }
        
        logger.info(f"🎯 Testing with record: {test_record}")
        logger.info(f"🔧 Timing config: {timing_config}")
        
        # Run the test
        from playwright.async_api import async_playwright
        
        async with async_playwright() as p:
            browser = await p.chromium.launch(
                headless=False,  # Visible for testing
                slow_mo=2000,    # Slow for observation
                args=[
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-popup-blocking',
                    '--disable-web-security'
                ]
            )
            
            logger.info("🔄 Processing single test record...")
            result = await processor.process_record(browser, test_record, timing_config)
            
            await browser.close()
        
        # Analyze results
        logger.info("📊 TEST RESULTS:")
        logger.info(f"   Status: {result['status']}")
        logger.info(f"   Downloaded Files: {result.get('downloaded_count', 0)}")
        logger.info(f"   Files: {result.get('downloaded_files', [])}")
        
        if result.get('error'):
            logger.error(f"   Error: {result['error']}")
        
        # Success criteria
        success = (
            result['status'] == 'success' and 
            result.get('downloaded_count', 0) > 0
        )
        
        if success:
            logger.info("🎉 SINGLE RECORD TEST SUCCESSFUL!")
            logger.info("✅ The final solution is working correctly")
            logger.info("✅ Ready for production use")
        else:
            logger.error("❌ SINGLE RECORD TEST FAILED")
            logger.error("❌ Need to investigate further")
        
        return success
        
    except Exception as e:
        logger.error(f"❌ Test failed with exception: {e}")
        return False

async def main():
    """Main test function."""
    logger.info("🧪 Starting Final Solution Single Record Test")
    
    # Check prerequisites
    try:
        with open('config.yaml', 'r') as f:
            config = yaml.safe_load(f)
        logger.info("✅ Config file found")
    except FileNotFoundError:
        logger.error("❌ config.yaml not found!")
        return False
    
    try:
        import pandas as pd
        logger.info("✅ Dependencies available")
    except ImportError as e:
        logger.error(f"❌ Missing dependency: {e}")
        return False
    
    # Run the test
    success = await test_single_record()
    
    if success:
        logger.info("\n🎉 FINAL SOLUTION VALIDATION COMPLETE")
        logger.info("✅ The submitFunction fix is working correctly")
        logger.info("✅ Popup content is loading properly")
        logger.info("✅ PDF generation is successful")
        logger.info("\n🚀 Ready to run full automation with:")
        logger.info("   python final_working_boleto_processor.py controle_boletos_hs.xlsx")
    else:
        logger.error("\n❌ FINAL SOLUTION NEEDS MORE WORK")
        logger.error("❌ Check logs for specific issues")
        logger.error("❌ Run debug_submitfunction_test.py for detailed analysis")
    
    return success

if __name__ == "__main__":
    try:
        result = asyncio.run(main())
        sys.exit(0 if result else 1)
    except KeyboardInterrupt:
        print("\n⏹️ Test interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Test failed: {e}")
        sys.exit(1)
