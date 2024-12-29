'''
calls functions to process the input excel file and generate a 
'''
import os
from flask import current_app
import zipfile

# import sys
# sys.path.append(os.path.join(os.path.dirname(__file__), "pantry_calls/make_lists"))
from caller_list_transform import make_guests_per_caller_lists, make_caller_pdfs, Caller_lists, get_fridays_date_string

import contextlib
@contextlib.contextmanager
# https://stackoverflow.com/questions/431684/how-do-i-cd-in-python
def pushd(new_dir):
    previous_dir = os.getcwd()
    os.chdir(new_dir)
    try:
        yield
    finally:
        os.chdir(previous_dir)

def run_script(input_file):
    # current_app.logger.info(f"input_file= {input_file}")

    Caller_lists = make_guests_per_caller_lists(input_file)

    processing_report_file = os.path.join(current_app.config['UPLOAD_FOLDER'], 'excel_processing_report.txt')
    status_str = ''
    if not Caller_lists.success:
        current_app.logger.warning(f"Failure: {Caller_lists.message}")
        status_str = f"Failure: {Caller_lists.message}"
    else:
        # date_str = get_fridays_date_string()
        success_list, failure_list = make_caller_pdfs(Caller_lists.caller_mapping_dict, Caller_lists.guest_dict, \
                        get_fridays_date_string(), out_pdf_dir=current_app.config['UPLOAD_FOLDER'])

        if len(Caller_lists.no_guest_list) > 0:
            status_str = "Callers with no guests: " + ', '.join(Caller_lists.no_guest_list) + "\n"
        else:
            status_str = "All callers have guests.\n"
        if "Do-Not-Call" in Caller_lists.caller_mapping_dict:
            # current_app.logger.info(f"Caller_lists.caller_mapping_dict['Do-Not-Call']= {Caller_lists.caller_mapping_dict['Do-Not-Call']}")
            # example:  Caller_lists.caller_mapping_dict['Do-Not-Call']= [['Guest6', None]]
            status_str += "Guests on the Do-Not-Call list: "
            for guest in Caller_lists.caller_mapping_dict["Do-Not-Call"]:
                status_str += f"{guest[0]} " # guest[0] is the guest UserName and guest[1] is the Caller's note
            status_str += "\n"
        # else:
        #     status_str += "No guests on Do-Not-Call list.\n"
        if len(success_list) > 0:
            status_str += "PDFs generated for: " + ', '.join(success_list) + "\n"
        if len(failure_list) > 0:
            status_str += "PDF generation failed for: " + ', '.join(failure_list)

    with open(processing_report_file, 'w') as f:
        f.write(status_str)

    for item in os.listdir(current_app.config['UPLOAD_FOLDER']):
        if item.endswith(".pdf"):
            current_app.logger.warning(f"will add {item}")
        if item.endswith(".txt"):
            current_app.logger.warning(f"will add {item}")

    return processing_report_file


'''
    with pushd(current_app.config['UPLOAD_FOLDER']):
        current_app.logger.warning(f"In {os.getcwd()}")
        with zipfile.ZipFile("call_lists", 'w') as zipf:
            for foldername, subfolders, filenames in os.walk("."):
                for filename in filenames:
                    if filename.endswith('.pdf') or filename.endswith('.txt'):
                        zipf.write(filename)

'''


def make_archive(source_dir, output_filename):
    with zipfile.ZipFile(output_filename, 'w') as zipf:
        for foldername, subfolders, filenames in os.walk(source_dir):
            for filename in filenames:
                if filename.endswith('.pdf') or filename.endswith('.txt'):
                    file_path = os.path.join(foldername, filename)
                    zipf.write(file_path, os.path.relpath(file_path, source_dir))

# Example usage
# source_directory = '/path/to/source_directory'
# output_zipfile = '/path/to/output.zip'
# make_archive(source_directory, output_zipfile)
