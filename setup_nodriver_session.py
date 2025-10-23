#!/usr/bin/env python3
"""
Setup NoDriver session for verificacion-correo.

This script creates a NoDriver-specific session file (nodriver_state.json)
that contains authentication cookies for OWA. This is separate from the
Playwright session (state.json) because NoDriver handles sessions differently.

Usage:
    python setup_nodriver_session.py

The script will:
1. Open Chrome with NoDriver (undetected mode)
2. Navigate to OWA login page
3. Wait for you to log in manually
4. Save the session to nodriver_state.json
"""

import asyncio
import json
from pathlib import Path

try:
    import nodriver as uc
except ImportError:
    print("❌ NoDriver not installed")
    print("   Install with: pip install nodriver")
    exit(1)


async def setup_session():
    """Interactive session setup for NoDriver."""
    print("="*70)
    print("NoDriver Session Setup for verificacion-correo")
    print("="*70)
    print()
    print("This will create a NoDriver-specific session file.")
    print("NoDriver uses Chrome with advanced anti-detection.")
    print()

    # Configuration
    owa_url = "https://correoweb.madrid.org/owa/#path=/mail"
    session_file = Path("nodriver_state.json")

    print(f"🌐 OWA URL: {owa_url}")
    print(f"💾 Session will be saved to: {session_file}")
    print()

    try:
        # Start NoDriver browser
        print("🚀 Starting NoDriver (Chrome with anti-detection)...")
        browser = await uc.start(
            headless=False,
            browser_args=[
                '--lang=es-ES',
                '--accept-lang=es-ES',
            ]
        )
        print("✅ Browser started")

        # Get main page
        page = browser.main_tab
        if not page:
            page = await browser.get("about:blank")

        # Navigate to OWA
        print(f"\n📧 Navigating to OWA...")
        await page.get(owa_url)
        await asyncio.sleep(2)

        print()
        print("="*70)
        print("⏸️  MANUAL LOGIN REQUIRED")
        print("="*70)
        print()
        print("The Chrome browser window should now be open.")
        print()
        print("Please complete these steps:")
        print("  1. Log in to OWA manually in the browser")
        print("  2. Wait until you see your inbox")
        print("  3. Come back here and press ENTER")
        print()
        print("The browser will stay open for up to 10 minutes.")
        print("You can close it manually if you finish earlier.")
        print()

        # Wait for user to log in (with timeout)
        try:
            # Wait for user input with timeout
            await asyncio.wait_for(
                asyncio.get_event_loop().run_in_executor(
                    None,
                    lambda: input("Press ENTER after logging in: ")
                ),
                timeout=600  # 10 minutes
            )
        except asyncio.TimeoutError:
            print("\n⏱️  Timeout reached (10 minutes)")

        # Check if we're authenticated
        current_url = page.url
        print(f"\n📍 Current URL: {current_url}")

        if 'login' in current_url.lower() or 'signin' in current_url.lower():
            print("\n❌ Still on login page")
            print("   Session was not saved. Please try again and complete the login.")
            return False

        print("✅ Appears to be authenticated")

        # Save cookies using NoDriver's native format
        print(f"\n💾 Saving session to {session_file}...")

        # Get all cookies
        try:
            import nodriver.cdp.network as cdp_network
            cookies_result = await page.send(cdp_network.get_all_cookies())
            cookies = cookies_result.cookies if hasattr(cookies_result, 'cookies') else []

            # Convert CDP cookies to JSON-serializable format
            session_data = {
                'cookies': [
                    {
                        'name': c.name,
                        'value': c.value,
                        'domain': c.domain,
                        'path': c.path,
                        'secure': c.secure if hasattr(c, 'secure') else False,
                        'httpOnly': c.http_only if hasattr(c, 'http_only') else False,
                        'sameSite': str(c.same_site) if hasattr(c, 'same_site') else 'None',
                        'expires': c.expires if hasattr(c, 'expires') else -1,
                    }
                    for c in cookies
                ],
                'url': current_url,
                'timestamp': str(asyncio.get_event_loop().time())
            }

            # Save to file
            with open(session_file, 'w', encoding='utf-8') as f:
                json.dump(session_data, f, indent=2)

            print(f"✅ Session saved successfully")
            print(f"   Cookies saved: {len(session_data['cookies'])}")
            print()
            print("="*70)
            print("✅ SETUP COMPLETE")
            print("="*70)
            print()
            print("You can now use NoDriver anti-detection mode with this session.")
            print("The session file will be used automatically when NoDriver is enabled.")
            print()

            return True

        except Exception as e:
            print(f"❌ Error saving cookies: {e}")
            import traceback
            traceback.print_exc()
            return False

    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        print("\n🧹 Closing browser...")
        try:
            await browser.stop()
        except:
            pass


if __name__ == "__main__":
    success = asyncio.run(setup_session())
    exit(0 if success else 1)
