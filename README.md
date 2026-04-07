# Parser mouser.com

A versatile parser for looking up electronic parts on **Mouser** by Part Number (both **Mouser PN** and **Manufacturer PN**) using the official Mouser Search API. It provides a CLI, an Interactive Console Menu, and a feature-rich Telegram Bot.

## Features

- **Search Capabilities**: Search by **Mouser Part Number** or **Manufacturer Part Number** (via keyword fallback).
- **Multiple Interfaces**:
  - **CLI (`mouser_cli.py`)**: Command-line execution with argument passing.
  - **Interactive Menu (`mouser_menu.py`)**: Console-based menu for executing queries and processing files.
  - **Telegram Bot (`telegram_bot.py`)**: Send PNs via text or attach an Excel `.xlsx` list to get fully formatted results directly in the chat.
- **Rich Data Extraction**: Fetches Mouser PN, Manufacturer PN, Manufacturer, Category (EN), Description, Image URL, Datasheet URL, and Compliance info.
- **Automated Translation**: Translates Categories and Descriptions from English to Russian using Google Translator.
- **Excel Batch Processing**: Upload an Excel file with a column named like `Part no.`, and the parser will automatically fetch data for all parts and return a fully detailed `.xlsx` file.
- **Local Database Cache (`SQLite`)**:
  - Prevents redundant requests by caching API responses locally to save Mouser API rate limits.
  - Tracks and stores user search history.
- **Smart Excel Highlighting**:
  - Highlights rows based on their specific product category:
    - 🟨 **Pale-yellow**: Categories requiring permissive documents.
    - 🟥 **Light-red**: Categories requiring mandatory labeling.
    - 🟦 **Light-blue**: Both documents and labeling required.
  - Highlights cells based on intellectual property (ТРОИС):
    - 🟪 **Bold Purple Font**: Brands present in the local Intellectual Property Registry.
- **API Key Rotation & Rate Limiting**: Supports multiple API keys (`MOUSER_API_KEY`, `MOUSER_API_KEYS`, `MOUSER_API_KEY_1`, etc.) to distribute request load and avoid 429 errors.

## Installation

### 1. Create and activate a virtual environment
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# Linux/macOS:
source .venv/bin/activate
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Environment Variables Setup
Copy the example configuration file:
```bash
cp .env.example .env
```
Edit the `.env` file to include your configurations:
- `MOUSER_API_KEYS` (or single `MOUSER_API_KEY`)
- `TELEGRAM_BOT_TOKEN`, `ADMIN_USER_ID`, `ALLOWED_THREAD_ID` (if you plan to use the Telegram Bot)

## Usage

### Interactive Menu
Run the interactive console menu:
```bash
python mouser_menu.py
```
Follow the prompts to enter one or multiple PNs, view the results table, read Excel files, and export data.

### CLI
Run the CLI for direct console outputs or scripts:
```bash
python mouser_cli.py MAX3232 579-24LC02B-I/SN --format xlsx --out report.xlsx
```
Use `python mouser_cli.py -h` to see all available arguments.

### Telegram Bot
Run the background worker to serve Telegram requests:
```bash
python telegram_bot.py
```
For regular use, simply send PNs as a message or attach an `.xlsx` file.
Use the `/menu` command to access an inline menu for managing search history, cache statistics, user authorization lists (Admin only), Intellectual Property lists, and highlighting categories.
