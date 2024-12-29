'''
calls functions to process the input excel file and generate a 
'''
import os
from flask import current_app
import zipfile
from caller_list_transform import make_guests_per_caller_lists, make_caller_pdfs, Caller_lists

def get_fridays_date_string():
   import datetime
   today = datetime.date.today()
   # print(f"{today.weekday()=}")
   # Sunday is 6, Monday is 0, Tuesday is 1, Wednesday is 2, Thursday is 3, Friday is 4, Saturday is 5
   # find the next Friday
   days_ahead = 4 - today.weekday()
   if days_ahead <= 0: # Target day already happened this week
      days_ahead += 7
   # print(f"{days_ahead=}")
   target_date = today + datetime.timedelta(days=days_ahead)
   return target_date.strftime('%Y-%m-%d')


def run_script(input_file):
    # current_app.logger.info(f"input_file= {input_file}")

    pantry_date_str = get_fridays_date_string()
    Caller_lists = make_guests_per_caller_lists(input_file)

    processing_report_filename = 'excel_processing_report.txt'
    processing_report_filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], f'{pantry_date_str}_{processing_report_filename}')
    zip_filename = 'call_lists.zip'
    zip_filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], f'{pantry_date_str}_{zip_filename}')
    status_str = ''
    if not Caller_lists.success:
        current_app.logger.warning(f"Failure: {Caller_lists.message}")
        status_str = f"Failure: {Caller_lists.message}"
    else:
        # remove any existing files in the upload folder
        for item in os.listdir(current_app.config['UPLOAD_FOLDER']):
            if item.endswith(".pdf") or item.endswith(".txt"):
                os.remove(os.path.join(current_app.config['UPLOAD_FOLDER'], item))
    
        success_list, failure_list = make_caller_pdfs(Caller_lists.caller_mapping_dict, Caller_lists.guest_dict, \
                        pantry_date_str, out_pdf_dir=current_app.config['UPLOAD_FOLDER'])

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

    with open(processing_report_filepath, 'w') as f:
        f.write(status_str)

    with zipfile.ZipFile(zip_filepath, 'w') as zipf:
        for item in os.listdir(current_app.config['UPLOAD_FOLDER']):
            if item.endswith(".pdf"):
                zipf.write(os.path.join(current_app.config['UPLOAD_FOLDER'], item), item)
            if item.endswith(".txt"):
                zipf.write(os.path.join(current_app.config['UPLOAD_FOLDER'], item), item)

    return zip_filepath
