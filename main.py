import os, csv, win32com.client
from datetime import datetime as dt
from utils import process_file

cwd = os.getcwd()
LOGS_FOLDER = os.path.join(cwd, 'Logs')

# Create Logs Folder
if not os.path.exists(LOGS_FOLDER):
    os.mkdir(LOGS_FOLDER)

LOGS_FILE = os.path.join(cwd, 'Logs', 'Logs.csv')
LOGS_HEADER = [
    'Filename',
    'Date Processed',
    'Time Processed',
    'Added to Excel',
    'Remarks'
]
EXCEL_PATH = os.path.join(cwd, 'Compiled_Surveys.xlsx')
EXCEL_HEADER = [
    'Name',
    'Date of Birth',
    'Contact',
    'Company',
    'Job Type',
    'MCQ_1',
    'MCQ_2',
    'MCQ_3',
    'MCQ_4',
    'MCQ_5',
    'FRQ_1',
    'FRQ_2',
    'FRQ_3',
    'Feedback',
    'Total MCQ Score'
]

# Start up Word and Excel
word = win32com.client.Dispatch('Word.Application')
excel = win32com.client.Dispatch('Excel.Application')

word.Visible = False
excel.Visible = False

# Create Logs File if does not exist
if not os.path.exists(LOGS_FILE):
    print('Creating Logs File...')
    with open(LOGS_FILE, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(LOGS_HEADER)

# Create Excel File if does not exist
if not os.path.exists(EXCEL_PATH):
    print('Creating Excel File...')
    wb = excel.Workbooks.Add()
    sheet = wb.ActiveSheet

    # Write data to cells
    for idx in range(len(EXCEL_HEADER)):
        sheet.Cells(1, idx + 1).Value = EXCEL_HEADER[idx]

    # Save the Workbook
    wb.SaveAs(EXCEL_PATH)

    wb.Close(SaveChanges=True)

files = [filename for filename in os.listdir(os.path.join(cwd, 'Unprocessed_Surveys')) if filename.endswith('.docx')]

# Track number of files processed
files_processed = 0
total = len(files)

print('Starting to process files...')

for filename in files:
    word_doc = os.path.join(cwd, 'Unprocessed_Surveys', filename)
    added_to_excel, remarks = process_file(word_doc=word_doc, word=word, excel_doc=EXCEL_PATH, excel=excel)

    # Add to the Logs File
    cur_date = dt.now().date()
    cur_time = dt.now().time()

    with open(LOGS_FILE, 'a', newline = '') as f:
        writer = csv.writer(f)
        writer.writerow([
            filename,
            cur_date,
            cur_time,
            added_to_excel,
            remarks
        ])

    if added_to_excel:
        os.rename(
            os.path.join(cwd, 'Unprocessed_Surveys', filename),
            os.path.join(cwd, 'Processed_Surveys', filename)
        )

    files_processed += 1
    print(f'Current Progress: {files_processed} / {total}')


# Close Word and Excel
word.Quit()
excel.Quit()

print('Done')