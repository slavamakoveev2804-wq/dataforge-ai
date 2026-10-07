# DataForge AI — MVP v0.2

A small web app that converts PDF tables/text into an Excel workbook.

## Run locally

Windows PowerShell:

    py -m venv .venv
    .\.venv\Scripts\Activate.ps1
    pip install -r requirements.txt
    python app.py

Then open:

    http://127.0.0.1:8000

## What works

- PDF upload
- PDF table extraction with pdfplumber
- Excel workbook output (.xlsx)
- Separate worksheet for each detected table
- Fallback text extraction when no table is detected
- Summary worksheet
- Drag & drop UI
- File-size limit
- Health endpoint at /health

## What is intentionally next

- OCR for scanned PDFs
- Secure paid conversion flow with Stripe
- Rate limiting / abuse protection
- User accounts and conversion history
- Better table reconstruction for difficult PDFs
- Production analytics

## Payment

Set STRIPE_PAYMENT_LINK to a Stripe Payment Link URL. The site then shows a real "Unlock for $5" button.

Do not collect card details yourself. Use Stripe-hosted Checkout/Payment Links.

## Public deployment

This repository is prepared for Render. Connect the repository to Render and create a Web Service using `render.yaml`, or deploy with the supplied Dockerfile. Set `STRIPE_PAYMENT_LINK` after creating a Stripe Payment Link.

Before charging real customers, test:
- normal text-based PDFs
- multi-page PDFs
- PDFs with multiple tables
- malformed/empty PDFs
- file-size rejection
- payment link
- privacy/terms pages

Do not claim OCR or perfect table reconstruction until those features are implemented and tested.
