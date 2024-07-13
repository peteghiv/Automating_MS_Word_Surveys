import os
import win32com
import win32com.client
from classes import Response, Form_Field

PERSONAL_INFO_FIELDS = [
    'Name',
    'DOB',
    'Mobile_Number',
    'Company',
    'Job_Type'
]

def get_form_fields(word_doc: str, word: win32com.client) -> list:
    '''
        Gets the form fields from a Word Document

        Args:
            - word_doc  :   File path to the required Word Document
            - word      :   win32com.client object for Word

        Return Vals:
            - A list of all the form fields.
    '''

    form_fields = []

    if not os.path.exists(word_doc):
        raise FileNotFoundError(f"The file {word_doc} does not exist.")
    
    try:
        doc = word.Documents.Open(word_doc)
        
        for field in doc.FormFields:
            if field.Type == 70:  # Text Field
                field_info = {
                    'name': field.Name,
                    'type': 'text',
                    'value': field.Result
                }
            elif field.Type == 71:  # Checkbox
                field_info = {
                    'name': field.Name,
                    'type': 'checkbox',
                    'value': field.CheckBox.Value
                }
            elif field.Type == 7:  # Dropdown
                field_info = {
                    'name': field.Name,
                    'type': 'dropdown',
                    'value': field.Result,
                    'options': [item.Text for item in field.DropDown.ListEntries]
                }
            else:
                field_info = {
                    'name': field.Name,
                    'type': 'unknown',
                    'value': field.Result
                }
            form_fields.append(field_info)
        
        doc.Close(False)  # Ensure proper closing
    except Exception as e:
        print(f"An error occurred: {e}")
    
    return form_fields


def process_file(word_doc: str, word: win32com.client, excel_doc: str, excel: win32com.client) -> tuple:
    '''
        Processes the response, and adds it to the Excel sheet if valid.
        Otherwise the sheet will not be updated

        Args:
            - word_doc  :   File path to the required Word Document
            - word      :   win32com.client object for Word
            - excel_doc :   File path to the required Excel Document
            - excel     :   win32com.client object for Excel

        Return Vals:
            - added_to_excel (bool)   :     Whether the response was stored in Excel
            - remarks        (str)    :     Additional Remarks
    '''
    form_fields = get_form_fields(word_doc, word)

    if len(form_fields) == 0:
        return None
    
    response = Response()

    for item in form_fields:
        field = Form_Field(item['name'], item['value'])

        if field.name in PERSONAL_INFO_FIELDS:
            response.Personal_Info.add_field(field)
        elif 'FRQ' in field.name:
            _, frq_sn = field.name.split('_')
            response.frq_arr[int(frq_sn) - 1].add_field(field)
        elif field.name == 'Add_Feedback':
            response.feedback.add_field(field)
        # If none of these, can only be the MCQ Likert Scale
        else:
            _, mcq_sn, _ = field.name.split('_')
            response.mcq_arr[int(mcq_sn) - 1].add_field(field)

    result = response.generate_report()

    if not result:
        # Get the error message
        _, msg = response.is_valid()
        return (False, msg)
    
    add_to_excel(response, excel, excel_doc)
    return (True, 'Valid Response')

def add_to_excel(response: Response, excel: win32com.client, excel_doc: str) -> bool:
    '''
        Adds the processed response into the Excel sheet

        Args:
            - response  :   Processed Response() object from process_file()
            - excel     :   win32com.client object for Excel
            - excel_doc :   File path to the required Excel Document
    '''
    error_flag = False
    wb = excel.Workbooks.Open(excel_doc)
    
    try: 
        sheet = wb.Sheets['Sheet1']

        # Get the row to insert data
        row_number = sheet.UsedRange.Rows.Count + 1
        idx = 1
        for item in response.report:
            sheet.Cells(row_number, idx).Value = str(response.report[item])
            idx += 1
    except Exception as e:
        print(f'Error: {e}')
        error_flag = True
    finally:
        wb.Close(SaveChanges=True)
    
    if error_flag:
        return False
    return True