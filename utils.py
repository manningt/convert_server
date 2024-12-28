'''
calls functions to process the input excel file and generate a 
'''
import os
from flask import current_app

# import os, sys
# sys.path.append(os.path.join(os.path.dirname(__file__), "pantry_calls/make_lists"))
from caller_list_transform import make_guests_per_caller_lists, make_caller_pdfs, Caller_lists, get_fridays_date_string


def run_script(input_file):
    # current_app.logger.info(f"input_file= {input_file}")

    Caller_lists = make_guests_per_caller_lists(input_file)

    if not Caller_lists.success:
        current_app.logger.warning(f"Failure: {Caller_lists.message}")
        failure_message_file = os.path.join(current_app.config['UPLOAD_FOLDER'], 'failure_message.txt')
        with open(failure_message_file, 'w') as f:
            f.write(f"Failure: {Caller_lists.message}")
        return failure_message_file
    else:
        callers_without_guests_file = os.path.join(current_app.config['UPLOAD_FOLDER'], 'callers_without_guests.txt')
        if len(Caller_lists.no_guest_list) > 0:
            status_str = f"Callers with no guests: {Caller_lists.no_guest_list}"
        else:
            status_str = "All callers have guests."
        with open(callers_without_guests_file, 'w') as f:
            f.write(status_str)

        date_str = get_fridays_date_string()
        make_caller_pdfs(Caller_lists.caller_mapping_dict, Caller_lists.guest_dict, date_str, out_pdf_dir=current_app.config['UPLOAD_FOLDER'])

    return callers_without_guests_file
