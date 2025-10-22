# save_session.py
from playwright.sync_api import sync_playwright
from config import PAGE_URL, BROWSER_CONFIG

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Navegar a la URL configurada
        page.goto(PAGE_URL)

        print("Abre el navegador y realiza el login manualmente.")
        input("Cuando termines, presiona ENTER en esta terminal para guardar la sesión...")

        storage_path = "state.json"
        context.storage_state(path=storage_path)
        print(f"Sesión guardada en {storage_path}")

        context.close()
        browser.close()

if __name__ == "__main__":
    main()