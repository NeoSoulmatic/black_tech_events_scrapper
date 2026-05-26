#!/bin/bash
# Black Tech Events Scraper v2 — setup

set -e

echo "=== Black Tech Events Scraper v2 — Setup ==="
echo ""

echo "Creating virtual environment (.venv)..."
python3 -m venv .venv

echo ""
echo "Installing Python dependencies..."
.venv/bin/pip install -r requirements.txt -q

echo ""
echo "Installing Playwright browser (Chromium)..."
.venv/bin/playwright install chromium

echo ""
echo "Creating .env from template..."
if [ ! -f .env ]; then
    cp .env.example .env
    echo "  Created .env — open it and fill in your API keys."
else
    echo "  .env already exists — skipping."
fi

echo ""
echo "============================================================"
echo "  Setup complete!"
echo "============================================================"
echo ""
echo "Next steps:"
echo ""
echo "  1. Open .env and add:"
echo "       GOOGLE_API_KEY  — from https://console.cloud.google.com"
echo "       GOOGLE_CSE_ID   — from https://programmablesearchengine.google.com"
echo "       (see .env.example for full step-by-step instructions)"
echo ""
echo "  2. Run the scraper:"
echo "       .venv/bin/python scraper.py"
echo ""
echo "  Output: black_tech_events_YYYYMMDD_HHMMSS.xlsx"
echo ""
echo "  Optional — push results to Google Sheets:"
echo "       Set ENABLE_GOOGLE_SHEETS=true in .env"
echo "       Add credentials.json (OAuth) or service_account.json"
echo ""
