# SheetDown-extract-structured-data-from-Excel

## Motivation
While working on RAG (Retrieval-Augmented Generation) systems, we found that most existing Excel parsing tools fail to properly extract non-text content (e.g., images), making them unsuitable for multimodal RAG applications. To address this, we developed this tool to comprehensively extract structured data (text and images) from Excel files, ensuring better compatibility with multimodal RAG pipelines.

## Method
Uses openpyxl to parse Excel files, extracting text and embedded images.
Outputs in Markdown format, with tables and merged cells rendered using HTML (Markdown-compatible).
Extracted images are saved in an images/ directory, and their paths are embedded in the corresponding cells using HTML <img> tags or Markdown <img> tags.
This method can process Excel files with multiple sheets.

## Requirements
Python ≥ 3.8

## Installation
```bash
pip install openpyxl
```
## Example
Modify the file path to your target input file.
```bash
python parser.py
```
Original table
![image](https://github.com/yuelang222/SheetDown-extract-structured-data-from-Excel/blob/main/example_raw.png)
Rendered Markdown output
![image](https://github.com/yuelang222/SheetDown-extract-structured-data-from-Excel/blob/main/example_extracted.png)
