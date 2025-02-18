'''
calls functions to process the input excel file and generate a 
'''
import os
from flask import current_app
import shutil
from caller_list_transform import make_guests_per_caller_lists, make_caller_pdfs, Caller_lists


def run_script(input_file):
    # current_app.logger.info(f"input_file= {input_file}")

    TEST_EXCEPTION = False
    if TEST_EXCEPTION:
        raise ValueError("Test")

    os.chdir(current_app.config['UPLOAD_FOLDER'])
    ZIP_DIR_PREFIX = "zip_dir_"
    CALL_LISTS_DIR_PREFIX = "caller_lists_"
    PDFS_DIR_PREFIX = "caller_PDFs_"

    ERROR_REPORT_FILENAME = 'error_report.txt'
    error_report_filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], f'{ERROR_REPORT_FILENAME}')

    # extract day info from filename, e.g. Master List for calling Jan 17 Pantry Day.xlsx
    START = 0
    END = 1
    PANTRY_DAY_DELIMITERS = ["calling ", ".xlsx"]
    pantry_day_index = [-1,-1]
    
    for index, element in enumerate(PANTRY_DAY_DELIMITERS):
        pantry_day_index[index] = input_file.lower().find(element)
        if pantry_day_index[index] == -1:
            with open(error_report_filepath, 'w') as f:
                f.write(f"Error: Could not find '{element}' in the filename: {os.path.basename(input_file)}")
            return error_report_filepath
    
    pantry_day_index[START] += len(PANTRY_DAY_DELIMITERS[START])
    pantry_date_str = input_file[pantry_day_index[START]:pantry_day_index[END]].replace(" ", "_") # get_fridays_date_string()

    if False:
        with open(error_report_filepath, 'w') as f:
            f.write(f"OK: '{os.path.basename(input_file)}' passed checks; date_string= {pantry_date_str}")
        return error_report_filepath

    Caller_lists = make_guests_per_caller_lists(input_file)
    status_str = ''
    if not Caller_lists.success:
        current_app.logger.warning(f"Failure: {Caller_lists.message}")
        with open(error_report_filepath, 'w') as f:
            f.write(f"Failure: {Caller_lists.message}")
        return error_report_filepath
    else:
        if False:
            print(f"{Caller_lists.caller_mapping_dict= }")
            print(f"{Caller_lists.guest_dict=}")
        # remove any existing files in the upload folder
        for item in os.listdir():
            if item.startswith(CALL_LISTS_DIR_PREFIX):
                os.remove(item)
                current_app.logger.info(f"Deleted: {item}")
            if item.startswith(ZIP_DIR_PREFIX):
                shutil.rmtree(item)
                current_app.logger.info(f"Deleted: {item}")

        # make a 2 level nested directory to be zipped, so that unzippiing will result in a directory
        directory_name_L1 = f'{ZIP_DIR_PREFIX}{pantry_date_str}'
        os.makedirs(directory_name_L1)
        directory_name_L2 = f'{CALL_LISTS_DIR_PREFIX}{pantry_date_str}'
        directory_path_L2 = os.path.join(directory_name_L1, directory_name_L2)
        os.makedirs(directory_path_L2)
        directory_name_PDFs = f'{PDFS_DIR_PREFIX}{pantry_date_str}'
        directory_path_PDFs = os.path.join(directory_path_L2, directory_name_PDFs)
        os.makedirs(directory_path_PDFs)
        # UPLOAD_FOLDER/zip_dir_date/caller_lists_date/caller_PDFs_date
     
        # Currently generating PDFs for every caller, to keep it simpler
        # filtered_callers_dict = filter_callers(Caller_lists.caller_mapping_dict)        
        # success_list, failure_list = make_caller_pdfs(filtered_callers_dict, Caller_lists.guest_dict, \

        success_list, failure_list = make_caller_pdfs(Caller_lists.caller_mapping_dict, Caller_lists.guest_dict, \
                        pantry_date_str, out_pdf_dir=directory_path_PDFs)

        if len(Caller_lists.invalid_usernames) > 0:
            status_str += "Guests with invalid usernames: " + ', '.join(Caller_lists.invalid_usernames) + "\n\n"
        if len(Caller_lists.invalid_caller_names) > 0:
            status_str += "Callers with invalid names: " + '; '.join(Caller_lists.invalid_caller_names) + "\n\n"
        if len(Caller_lists.no_guest_list) > 0:
            status_str = "Callers with no guests: " + ', '.join(Caller_lists.no_guest_list) + "\n\n"
        if len(Caller_lists.guests_without_caller) > 0:
            status_str += "These guests don't have a caller: " + ', '.join(Caller_lists.guests_without_caller) + "\n\n"        
        if len(success_list) > 0:
            status_str += f"generated {len(success_list)} PDFs\n"
        if len(failure_list) > 0:
            status_str += "PDF generation failed for: " + ', '.join(failure_list)

    processing_report_filename = f'excel_processing_report_{pantry_date_str}.txt'
    processing_report_filepath = os.path.join(directory_path_L2, f'{processing_report_filename}')
    with open(processing_report_filepath, 'w') as f:
        f.write(status_str)

    zip_filename = f'call_lists_{pantry_date_str}'
    # change working directory to have zip file in the UPLOAD_FOLDER
    zip_filepath = shutil.make_archive(zip_filename, 'zip', directory_name_L1)

    current_app.logger.info(f"zip file created: '{zip_filepath}'") 
    return zip_filepath
