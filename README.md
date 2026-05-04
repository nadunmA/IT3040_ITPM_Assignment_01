# Assignment 1 - Singlish to Sinhala Transliteration Testing

## Overview

This project automates testing of the Chat Sinhala transliteration function for IT3040 - ITPM Assignment 1 using Playwright.
The objective is to evaluate how accurately chat-style Singlish inputs are converted into Sinhala output.

## GitHub Repository

Repository Link:
https://github.com/nadunmA/IT3040_ITPM_Assignment_01.git

## Note on Script Adaptation

## Note on Script Adaptation

The script was designed and executed using the official Chat Sinhala transliteration application:
https://www.pixelssuite.com/chat-translator.

## Project Structure

IT23192850/

- test_automation/
  - README.md
  - test_automation.py

- github_link.txt
- IT23192850.xlsx

## Prerequisites

- Python 3.11 or 3.12
- Google Chrome

## Clone the Repository

git clone https://github.com/nadunmA/IT3040_ITPM_Assignment_01.git
cd IT3040_ITPM_Assignment_01/test_automation

## Installation

pip install -U pip
pip install playwright openpyxl
playwright install

## Running the Tests

Navigate to the test_automation folder:
cd D:\IT23192850\test_automation

Run the automation script:
python test_automation.py --excel "..\IT23192850.xlsx" --input-col "Input" --expected-col "Expected output" --actual-col "Actual output" --status-col "Status" --wait-ms 15000 --retries 10 --retry-wait-ms 2000 --save-every 1

## Output

After execution, the script updates the Excel file with:

- Actual Output
- Status (PASS / FAIL)

## Submitted Files

The submission includes:

- Playwright automation project
- Completed Excel file (IT23192850.xlsx)
- GitHub repository link (github_link.txt)
- README file with setup and execution instructions
