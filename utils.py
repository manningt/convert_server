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
    return processing_report_file
