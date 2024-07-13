class Form_Field:
    def __init__(self, name: str = '', value: str = '') -> None:
        self.name = name
        self.value = value

class Likert_Scale:
    def __init__(self, SN: int = 0):
        self.SN = SN
        self.arr = []
        self.score = 0

    def is_valid(self) -> tuple:
        # Check that arr has 6 distinct Form Elements
        if len(self.arr) != 6:
            return (False, f'MCQ {self.SN}: Does not have 6 distinct Form Elements')
        
        temp = [False, False, False, False, False, False]
        no_checked = 0
        for field in self.arr:
            if field.value == True:
                no_checked += 1

            field_score = int(field.name.split('_')[-1])
            temp[field_score - 1] = True

        for item in temp:
            if item == False: # Means that this field is not present in self.arr
                return (False, f'MCQ {self.SN}: Does not have 6 distinct Form Elements')
            
        # Check that only 1 checkbox is ticked
        if no_checked != 1:
            if no_checked == 0:
                return (False, f'MCQ {self.SN}: No checkbox ticked')
            else:
                return (False, f'MCQ {self.SN}: More than 1 checkbox ticked')
            
        # Passes all tests
        return (True, f'MCQ {self.SN}: Valid')
    
    def generate_score(self) -> bool:
        # Check if all form elements present first
        valid, _ = self.is_valid()

        if not valid:
            return False
        
        # Look for the marked checkbox
        for field in self.arr:
            if field.value == True:
                field_score = int(field.name.split('_')[-1])
                self.score = field_score
                return True
        
        return False
    
    def add_field(self, new_field: Form_Field) -> bool:
        self.arr.append(new_field)

        valid, _ = self.is_valid()
        if valid:
            self.generate_score()

        return True
    
class Free_Response:
    def __init__(self, SN: int = 0) -> None:
        self.SN = SN
        self.response = ''

    def is_valid(self) -> tuple:
        # Check that response is not ''
        answer = (self.response != '')

        if answer:
            return (True, f'FRQ {self.SN}: Valid')
        else:
            return (False, f'FRQ {self.SN}: Missing Response')
    
    def add_field(self, new_field: Form_Field) -> bool:
        self.response = new_field.value.strip()

class Additional_Feedback:
    def __init__(self) -> None:
        self.response = ''

    def add_field(self, new_field: Form_Field) -> bool:
        self.response = new_field.value.strip()

class Personal_Info:
    def __init__(self) -> None:
        self.name = ''
        self.dob = ''
        self.contact = ''
        self.company = ''
        self.job_type = ''

    def add_field(self, new_field: Form_Field) -> bool:
        if new_field.name == 'Name':
            self.name = new_field.value.strip()
            return True
        if new_field.name == 'DOB':
            self.dob = new_field.value.strip()
            return True
        if new_field.name == 'Mobile_Number':
            self.contact = new_field.value.strip()
            return True
        if new_field.name == 'Company':
            self.company = new_field.value.strip()
            return True
        if new_field.name == 'Job_Type':
            if new_field.value == 'Please select Job Type':
                return False
            else:
                self.job_type = new_field.value
                return True
        
        # Invalid field
        return False
    
    def is_valid(self) -> tuple:
        missing_items = []
        
        # Check that all of the attributes are not empty
        if self.name == '':
            missing_items.append('Name')
        if self.dob == '':
            missing_items.append('Date of Birth')
        if self.contact == '':
            missing_items.append('Contact Number')
        if self.company == '':
            missing_items.append('Company')
        if self.job_type == '':
            missing_items.append('Job Type')
        
        if len(missing_items) > 0:
            return (False, f'Missing: {", ".join(missing_items)}')
        # If all attributes not empty, considered valid.
        return (True, 'Valid')
    
class Response:
    def __init__(self) -> None:
        self.Personal_Info = Personal_Info()
        self.mcq_arr = [Likert_Scale(i+1) for i in range(5)]
        self.frq_arr = [Free_Response(i+1) for i in range(3)]
        self.feedback = Additional_Feedback()
        # --------------------------------------
        self.report = {
            'Name'    : '',
            'DOB'     : '',
            'Contact' : '',
            'Company' : '',
            'Job Type': '',
            'MCQ_1'   : '',
            'MCQ_2'   : '',
            'MCQ_3'   : '',
            'MCQ_4'   : '',
            'MCQ_5'   : '',
            'FRQ_1'   : '',
            'FRQ_2'   : '',
            'FRQ_3'   : '',
            'Feedback': ''
        }

    def is_valid(self) -> tuple:
        # Check using is_valid function of each attribute
        message = []

        PI_valid, PI_msg = self.Personal_Info.is_valid()

        if not PI_valid:
            message.append(PI_msg)
        
        for mcq in self.mcq_arr:
            MCQ_valid, MCQ_msg = mcq.is_valid()
            if not MCQ_valid:
                message.append(MCQ_msg)
            
        for frq in self.frq_arr:
            FRQ_valid, FRQ_msg = frq.is_valid()
            if not FRQ_valid:
                message.append(FRQ_msg)

        if len(message) > 0:
            return (False, '; '.join(message))
            
        # Note: Feedback field is optional and therefore valid by default
        return (True, 'Valid')
    
    def generate_report(self) -> bool:
        valid, _ = self.is_valid()

        if not valid:
            return False
        
        self.report['Name'] = self.Personal_Info.name
        self.report['DOB'] = self.Personal_Info.dob
        self.report['Contact'] = self.Personal_Info.contact
        self.report['Company'] = self.Personal_Info.company
        self.report['Job Type'] = self.Personal_Info.job_type

        for i in range(len(self.mcq_arr)):
            self.report[f'MCQ_{i+1}'] = self.mcq_arr[i].score

        for i in range(len(self.frq_arr)):
            self.report[f'FRQ_{i+1}'] = self.frq_arr[i].response

        self.report['Feedback'] = self.feedback.response

        total = 0
        for i in range(len(self.mcq_arr)):
            total += self.mcq_arr[i].score
        self.report['MCQ Score'] = total

        return True