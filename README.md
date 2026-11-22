# SATO SA4 Excel Continuous Label Printer

A lightweight Windows desktop utility in Python for batch-printing multi-row stickers from Excel spreadsheets to a **SATO SA4** thermal label printer without paid label design software.

## Features
- Load `.xlsx` / `.xls` files and choose any worksheet.
- Map up to **4 lines per sticker** from Excel columns or static/custom text.
- Customize **font family, font size, bold, and italic** per line when using **Windows Driver (Font/Style)** mode, or send **Raw SBPL** commands.
- Configure sticker width/height (`mm`), copy count, and row range for continuous printing.

## Setup
```bash
pip install -r requirements.txt
python app.py