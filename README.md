# DOCX Migration V7

Enterprise-style DOCX migration framework.

## Features

- DOCX Analysis
- Paragraph Extraction
- Table Extraction
- Table Preservation
- Round-Trip Testing
- Validation

## Project Structure

docx-migration/

- docx_migration/ - core package (models, extractor, builder, validator, branding, template_manager, style_engine)
- config/ - document type rules, style maps, formatting/acceptance policy
- templates/ - production templates + reference/ source templates
- scripts/ - standalone dev tools (analyser.py, compare.py)
- tests/ - test_roundtrip.py
- samples/ - sample input/output documents
- main_migration.py - CLI entry point

## Installation

pip install -r requirements.txt

## Run Migration

python main_migration.py --source samples/file.docx --output samples/output.docx --document-type datasheet

## Analyse Document

python scripts/analyser.py samples/file.docx

## Round Trip Test

python tests/test_roundtrip.py

## Current Status

✅ Paragraph Preservation

✅ Table Preservation

✅ Object Inventory

⏳ Image Position Preservation

⏳ Headers & Footers

⏳ Textboxes