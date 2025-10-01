import re
from collections import defaultdict

def parse_fm_ch_from_key(key_string):
    parts = key_string.split(' - ')
    if len(parts) >= 2:
        try:
            credit_hour = float(parts[-1])
            full_marks = float(parts[-2])
            return full_marks, credit_hour
        except (ValueError, IndexError):
            return None, None
    return None, None
def map_marksheet_data_by_section(data):
    """Maps flat marksheet data into a structured dictionary based on user-defined sections."""
    mapped_data = {
        "student_info": {},
        "first_term": {"subjects": defaultdict(dict), "overall": {}},
        "second_term": {"subjects": defaultdict(dict), "overall": {}},
        "annual_internal_evaluation": {"subjects": defaultdict(dict)},
        "annual": {"subjects": defaultdict(dict), "overall": {}}
    }
    
    sorted_keys = sorted((k for k in data if re.match(r'^\d+', k)), key=lambda k: int(re.match(r'^\d+', k).group()))
    last_mark_type = None
    for key in sorted_keys:
        value = data[key]
        key_num = int(re.match(r'^\d+', key).group())
        clean_key = re.sub(r'^\d+\s*-\s*', '', key).strip()
        if 1 <= key_num <= 11:
            extract_personal_info(mapped_data, key, value)
        elif 12 <= key_num <= 37:
            extract_term(mapped_data, value, clean_key,'first_term')

        elif 38 <= key_num <= 62:
            extract_term(mapped_data, value, clean_key,'second_term')
        elif key_num >= 99:
          last_mark_type = final_evaluation(mapped_data, value, clean_key, last_mark_type)

    for section in ['first_term', 'second_term', 'annual']:
        if 'subjects' in mapped_data[section]:
             mapped_data[section]['subjects'] = dict(mapped_data[section]['subjects'])
    mapped_data['annual_internal_evaluation']['subjects'] = dict(mapped_data['annual_internal_evaluation']['subjects'])

    return mapped_data

def final_evaluation(mapped_data, value, clean_key,last_mark_type =None) -> str|None:
    current_subject  = extract_subject(clean_key)
    clean_key = clean_key.split(' - ')
    if current_subject:
        fm, ch = parse_fm_ch_from_key(' - '.join(clean_key)) 
        if 'Int. Mark' in clean_key:
            entry = mapped_data['annual']['subjects'][current_subject].setdefault('internal', {})
            entry['mark'] = value
            if fm is not None: entry['full_marks'] = fm
            if ch is not None: entry['credit_hour'] = ch
            last_mark_type = 'internal'

        elif 'Ext.Mark' in clean_key:
            entry = mapped_data['annual']['subjects'][current_subject].setdefault('external', {})
            entry['mark'] = value
            if fm is not None: entry['full_marks'] = fm
            if ch is not None: entry['credit_hour'] = ch
            last_mark_type = 'external'

        elif 'GP' in clean_key and last_mark_type:
            mapped_data['annual']['subjects'][current_subject][last_mark_type]['gp'] = value
        elif 'WGP' in clean_key and last_mark_type:
            mapped_data['annual']['subjects'][current_subject][last_mark_type]['wgp'] = value
        elif 'Total Mark' in clean_key:
            mapped_data['annual']['subjects'][current_subject]['total_mark'] = value
        elif 'Final GP' in clean_key: 
            mapped_data['annual']['subjects'][current_subject]['final_gp'] = value
        elif 'FinalGrade' in clean_key:
            mapped_data['annual']['subjects'][current_subject]['final_grade'] = value
    elif 'Grand Total' in clean_key:
        mapped_data['annual']['overall']['grand_total'] = value
    elif 'GPA' in clean_key:
        mapped_data['annual']['overall']['gpa'] = value
    elif 'Rank' in clean_key:
        mapped_data['annual']['overall']['rank'] = value
    return last_mark_type

def extract_subject(clean_key: str):
    clean = clean_key.strip()
    if '-' in clean:
        return clean.split('-')[0].strip()
    subject_map = {'Mathmatics': 'Mathematics', 'HPA': 'HPE', 'Local Sub': 'Local Subject'}
    all_subjects = list(subject_map.keys()) + ['English', 'Nepali', 'Science', 'Social', 'Local Subject', 'Opt.I','Opt.II']
    if(clean in all_subjects):
        return clean
    else:
        return None


def extract_term(mapped_data, value, clean_key, term_val):
    current_subject = extract_subject(clean_key)
    if(current_subject):
        if ' - Mark - ' in clean_key:
            mapped_data[term_val]['subjects'][current_subject]['mark'] = value
            fm, ch = parse_fm_ch_from_key(clean_key) 
            if fm is not None: 
                mapped_data[term_val]['subjects'][current_subject]['full_marks'] = fm
            if ch is not None:
                mapped_data[term_val]['subjects'][current_subject]['credit_hour'] = ch
        elif 'Grade' in clean_key:
            mapped_data[term_val]['subjects'][current_subject]['grade'] = value
        elif 'Gp' in clean_key:
            mapped_data[term_val]['subjects'][current_subject]['gp'] = value
    elif 'Total' in clean_key:
        mapped_data[term_val]['overall']['total'] = value
    elif 'GPA' in clean_key:
        mapped_data[term_val]['overall']['gpa'] = value
    elif 'Rank' in clean_key:
        mapped_data[term_val]['overall']['rank'] = value



def extract_personal_info(mapped_data, key, value):
    if 'S.N.' in key:
        mapped_data['student_info']['serial_number'] = value
    elif 'Roll No.' in key:
        mapped_data['student_info']['roll_number'] = value
    elif 'Name of the Student' in key:
        mapped_data['student_info']['name'] = value
    elif 'Gender' in key:
        mapped_data['student_info']['gender'] = value
    elif 'Caste' in key:
        mapped_data['student_info']['caste'] = value
    elif 'DOB' in key:
        mapped_data['student_info']['dob'] = value
    elif "Father" in key:
        mapped_data['student_info']['fathers_name'] = value
    elif "Mother" in key:
        mapped_data['student_info']['mothers_name'] = value
    elif "Parent" in key:
        mapped_data['student_info']['parents_name'] = value
    elif 'Address' in key:
        mapped_data['student_info']['address'] = value
    elif 'Contact No.' in key:
        mapped_data['student_info']['contact_number'] = value

