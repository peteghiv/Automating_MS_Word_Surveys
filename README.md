# Automating MS Word Surveys

This uses PyWin32 to process a feedback survey / form created in Microsoft Word.

## Creating the MS Word Survey
1. Use Legacy Form Fields in MS Word (Developer Tab)
2. Restrict Editing to only allow filling in of forms

## Usage
1. Place the filled up items in the subfolder "Unprocessed_Surveys"
2. Run main.py
3. Valid survey responses are collated in an Excel sheet, and the respective files are transferred to "Processed_Surveys"
4. Invalid survey responses remain in "Unprocessed_Surveys"
5. A Logs folder and CSV file are created to show any relevant validation errors in the files