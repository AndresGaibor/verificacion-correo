# save_session.py
from playwright.sync_api import sync_playwright

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Cambia por la URL de login que uses
        page.goto("https://correoweb.madrid.org/owa/#path=/mail")

        print("Abre el navegador y realiza el login manualmente.")
        input("Cuando termines, presiona ENTER en esta terminal para guardar la sesión...")

        storage_path = "state.json"
        context.storage_state(path=storage_path)
        print(f"Sesión guardada en {storage_path}")

        context.close()
        browser.close()

if __name__ == "__main__":
    main()