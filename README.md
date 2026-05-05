# IT23693272 - ITPM Assignment 1

## Assignment Title
Singlish to Sinhala Transliteration Testing using Playwright Automation

## Description
This project evaluates the accuracy of the Chat Sinhala transliteration tool available at:

https://www.pixelssuite.com/chat-translator

The purpose of this assignment is to test how accurately the system converts informal chat-style Singlish input into Sinhala output. The test cases are designed as negative test cases to identify transliteration inaccuracies and system limitations.

## Student Details
- Student ID: IT23693272
- Name: Yasuththara B G P
- Module: IT3040 - ITPM
- Assignment: Assignment 1 - Option 1
- Degree Program: BSc (Hons) in Information Technology
- Academic Year: Year 3

## Files Included
Assignment 1 - Test cases.xlsx  
test_automation.py  
README.md  
repo_link.txt  

## Requirements
- Python 3.10 or above
- Google Chrome
- Playwright
- openpyxl

## Installation
pip install playwright openpyxl  
python -m playwright install  

## How to Run
python test_automation.py --excel "Assignment 1 - Test cases.xlsx"

## Test Case Details
- Total test cases: 50
- Test type: Negative testing
- TC ID format: Neg_0001 to Neg_0050

## Output
After running the script, the Excel file will be updated with:
- Actual output
- Status

Most test cases result in FAIL because they are negative test cases designed to evaluate system limitations.

## Repository Link
https://github.com/poornigamage2003/IT23693272-ITPM-Assignment-1
