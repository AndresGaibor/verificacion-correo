"""
Test script to verify if NoDriver can extract the contact name from OWA.

This script uses NoDriver (undetected Chrome) to test if Microsoft OWA
shows the real name when automation is better hidden.
"""

import asyncio
from verificacion_correo.core.config import get_config
from verificacion_correo.core.antidetection import NoDriverManager, NoDriverConfig, UserAgentConfig


async def test_name_extraction():
    """Test name extraction with NoDriver."""
    config = get_config()

    # Setup NoDriver
    nodriver_config = NoDriverConfig(
        headless=False,
        sandbox=True,
        lang="es-ES"
    )
    ua_config = UserAgentConfig(
        rotate=True,
        pool_size=10,
        prefer_platform=None
    )

    nodriver_manager = NoDriverManager(nodriver_config, ua_config)

    try:
        # Start NoDriver browser
        print("🚀 Starting NoDriver (undetected Chrome)...")
        await nodriver_manager.start(session_file=str(config.get_session_file_path()))
        page = await nodriver_manager.get_page()
        print("✅ NoDriver started successfully")

        # Navigate to OWA
        print(f"📧 Navigating to {config.page_url}...")
        await page.get(config.page_url)
        print("⏳ Waiting for page to load completely...")
        await asyncio.sleep(8)  # Give more time for OWA to load

        # Check if we're authenticated
        current_url = page.url
        print(f"📍 Current URL: {current_url}")

        if 'login' in current_url.lower() or 'signin' in current_url.lower():
            print("❌ Not authenticated - session expired or invalid")
            print("   Please run: python copiar_sesion.py to create a new session")
            return

        print("✅ Page loaded and authenticated")

        # Wait for new message button to be available
        print("🔍 Looking for new message button...")
        new_msg_btn = None
        for attempt in range(10):
            new_msg_btn = await page.select(config.selectors.new_message_btn)
            if new_msg_btn:
                print(f"✅ Found new message button (attempt {attempt + 1})")
                break
            await asyncio.sleep(1)

        if not new_msg_btn:
            print("❌ New message button not found after 10 attempts")
            print("   The page may not have loaded correctly")
            print("   Taking screenshot for debugging...")
            await page.save_screenshot('debug_nodriver_page.png')
            print("   Screenshot saved to: debug_nodriver_page.png")
            return

        # Click new message button
        print("✉️  Opening new message...")
        await new_msg_btn.click()
        await asyncio.sleep(3)
        print("✅ New message opened")

        # Add test email to To field
        test_email = "AGM564@MADRID.ORG"
        print(f"📝 Adding email: {test_email}...")

        # Find To field
        input_boxes = await page.select_all('[role="textbox"]')
        to_field = None
        for box in input_boxes:
            name = await box.get_attribute('name')
            if name and config.selectors.to_field_name.lower() in name.lower():
                to_field = box
                break

        if to_field:
            await to_field.send_keys(test_email)
            await asyncio.sleep(3)

            # Blur field to create token
            await page.evaluate("document.activeElement.blur()")
            await asyncio.sleep(1)
            print("✅ Email added and tokenized")

            # Find and click the email token
            print("🖱️  Clicking on email token...")
            all_spans = await page.select_all('span')
            email_span = None
            for span in all_spans:
                try:
                    text = await span.text_content()
                    if text and text.strip().lower() == test_email.lower():
                        email_span = span
                        break
                except:
                    continue

            if email_span:
                await email_span.click()
                await asyncio.sleep(2)
                print("✅ Token clicked, popup should be visible")

                # Extract popup content
                print("📋 Extracting popup content...")
                popup = await page.select(config.selectors.popup)
                if popup:
                    popup_text = await popup.text_content()
                    print("\n" + "="*60)
                    print("POPUP TEXT CONTENT:")
                    print("="*60)
                    print(popup_text[:800])
                    print("="*60)

                    # Check if name appears
                    if "GOMEZ" in popup_text or "ALMUDENA" in popup_text:
                        print("\n✅ SUCCESS: Real name found in popup!")
                        print("   NoDriver successfully bypassed OWA anti-scraping")
                    else:
                        print("\n❌ FAILED: Real name NOT found in popup")
                        print("   OWA is still hiding the name even with NoDriver")
                        print(f"   Token email appears: {'AGM564@MADRID.ORG' in popup_text}")
                else:
                    print("❌ Popup not found")
            else:
                print("❌ Email token not found")
        else:
            print("❌ To field not found")

        # Wait to observe
        print("\n⏸️  Pausing for 10 seconds to observe...")
        await asyncio.sleep(10)

    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

    finally:
        print("\n🧹 Cleaning up...")
        await nodriver_manager.close()
        print("✅ Done")


if __name__ == "__main__":
    print("="*60)
    print("NoDriver Name Extraction Test")
    print("="*60)
    print()
    asyncio.run(test_name_extraction())
