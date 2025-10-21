# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python automation tool that uses Playwright to interact with the Madrid city webmail interface (correoweb.madrid.org/owa). The application automates the process of extracting contact information from email recipients by interacting with the OWA (Outlook Web Access) interface.

## Architecture

The project consists of two main scripts:

1. **copiar_sesion.py** - Session Management
   - Launches a browser for manual authentication
   - Saves authentication state to `state.json` for reuse
   - Must be run first to establish a valid session

2. **app.py** - Main Automation Script
   - Loads saved session from `state.json`
   - Navigates to OWA webmail interface
   - Automates recipient token interaction
   - Extracts contact information from popup cards using multiple strategies:
     - Regex pattern matching for emails, phones, postal codes, SIP addresses, and names
     - Specific DOM selectors when available
     - Text content fallback extraction

### Key Components in app.py

- **leer_correos_excel()**: Reads email addresses from an Excel file (`data/correos.xlsx`)
  - Reads first column (A) starting from row 2 (row 1 is header)
  - Returns list of email addresses
  - Falls back to default emails if file not found or empty
- **Configuration Constants** (`PAGE_URL`, `CORREO`): Define target URL and email addresses to process
  - `CORREO` is now dynamically loaded from Excel via `leer_correos_excel()`
- **Regex Patterns**: Pre-compiled patterns for extracting structured data (email, phone, postal address, SIP, name)
- **extract_from_popup_text()**: Regex-based extraction from raw popup text
- **popup_info()**: Main extraction function that:
  1. Waits for popup with selector `div._pe_Y[ispopup="1"]`
  2. Tries specific DOM selectors (e.g., `span._pe_c1._pe_t1` for name)
  3. Falls back to regex extraction from full text
  4. Consolidates results prioritizing specific selectors over regex matches
- **main()**: Orchestrates the automation workflow

## Development Setup

1. **Create virtual environment**:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

2. **Install dependencies**:
   ```bash
   pip install playwright openpyxl
   playwright install chromium
   ```

3. **Configure email addresses**:
   - Create or edit `data/correos.xlsx`
   - Place email addresses in column A, starting from row 2
   - Row 1 should contain the header "Correo"
   - Example structure:
     ```
     | A                 |
     |-------------------|
     | Correo            | <- Row 1 (header)
     | user1@madrid.org  | <- Row 2
     | user2@madrid.org  | <- Row 3
     | ...               |
     ```
   - An example file is provided with default emails

4. **Create session file**:
   ```bash
   python copiar_sesion.py
   # Follow the manual login prompts, then press ENTER to save session
   ```

5. **Run the automation**:
   ```bash
   python app.py
   ```

## Important Implementation Details

- **Session Persistence**: The `state.json` file contains authentication cookies and must exist before running `app.py`
- **DOM Selectors**: The script targets specific OWA interface elements:
  - New message button: `button[title="Escribir un mensaje nuevo (N)"]`
  - To field: `textbox` with role and name "Para"
  - Email tokens: Located via regex pattern matching on visible text
  - Popup cards: `div._pe_Y[ispopup="1"]`
- **Timing Strategy**: Uses explicit waits (`wait_for_timeout`, `wait_for_selector`) to handle dynamic content loading
- **Error Handling**: Uses try-except blocks for optional extractions (see `safe_text()` and popup selector waits)
- **Contact Extraction Strategy**:
  - Primary: Use specific DOM selectors (class names, autoids)
  - Fallback: Extract from raw text using regex
  - Name format expected: "APELLIDO, NOMBRE" (surname, firstname)
  - Office/job title: Heuristic looks for all-uppercase lines

## Common Patterns

- **Launching browser with saved session**:
  ```python
  browser = p.chromium.launch(headless=False)
  context = browser.new_context(storage_state="state.json")
  ```

- **Safe element extraction**:
  ```python
  def safe_text(locator):
      try:
          return locator.inner_text(timeout=1000).strip()
      except:
          return None
  ```

- **Filtering locators by email list**:
  ```python
  escaped_emails = [re.escape(email) for email in emails]
  pattern = re.compile(r"^(?:" + "|".join(escaped_emails) + r")$", re.IGNORECASE)
  email_tokens = email_tokens.filter(has_text=pattern)
  ```

## Key Dependencies

- **playwright**: Browser automation framework (requires `playwright install` after pip install)
- **openpyxl**: Library for reading/writing Excel files (.xlsx format)
- Python 3.13 (as indicated by venv structure)

## Files to Ignore

- `state.json`: Contains session authentication data (already in .gitignore)
- `.venv/`: Virtual environment directory
- `data/`: Directory for storing data files including:
  - `correos.xlsx`: Excel file containing email addresses to process (example file provided)
