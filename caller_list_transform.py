#!/usr/bin/env python3
'''
2 functions:
1. make_guests_per_caller_lists(in_filename) -> Caller_lists
2. make_caller_pdfs(caller_mapping_dict, guest_dict, date_str)

The input file is an Excel file with 3 sheets: 'guest-to-caller', 'callers', 'guests'
See class definition below for the NamedTuple Caller_lists returned by make_guests_per_caller_lists(in_filename)

The output is a PDF file for each caller with a list of guests for the next Friday.

'''

import sys, os
try:
   from openpyxl import load_workbook
   from fpdf import FPDF
except Exception as e:
   print(e)
   sys.exit(1)

import argparse
from pathlib import Path
from flask import current_app

from typing import NamedTuple  #not to be confused with namedtuple in collections
class Caller_lists(NamedTuple):
    success: bool = False
    message: str = ''
    caller_mapping_dict: dict = {}
    no_guest_list: list = []
    guest_dict: dict = {}
    invalid_usernames: list = []


def make_guests_per_caller_lists(in_filename):
   # returns the tuple Caller_lists
   GUESTS_SHEET_NAME = 'Master List'
   GUEST_FIRSTNAME = 1
   GUEST_LASTNAME = 2
   GUEST_USERNAME = 3
   GUEST_PASSWORD = 4
   GUEST_TOWN = 8
   GUEST_PHONE = 9
   GUEST_NOTES = 10

   Caller_lists()
   Caller_lists.success = False # default value didn't seem to work

   try:
      workbook = load_workbook(in_filename, data_only=True)
   except Exception as e:
      Caller_lists.message = f"Error when reading {in_filename}: {e}"
      return Caller_lists
   
   sheetnames = workbook.sheetnames
   # print(f"{workbook.sheetnames=}")

   # check that expected sheets are in the excel spreadsheet (other sheets are ignored)
   EXPECTED_SHEETNAMES = {'guest-to-caller', 'callers', GUESTS_SHEET_NAME}
   sheetnames_set = set(sheetnames)
   if EXPECTED_SHEETNAMES != sheetnames_set.intersection(EXPECTED_SHEETNAMES):
      Caller_lists.message = f"Error: expected '{EXPECTED_SHEETNAMES}' sheet names; found '{sheetnames_set}' in file '{in_filename}'"
      return Caller_lists

   #make dictionary of caller, [guests]
   mapping_dict = {}
   is_header = True
   for row in workbook['callers'].rows:
      # skip header row; there is only 1 column: the list of callers
      if is_header:
         is_header = False
      else:
         mapping_dict[row[0].value] = []
   
   is_header = True
   for row in workbook['guest-to-caller'].rows:
      # columns: 0=Guest, 1=Caller, 2=Change, 3=Note, 4=Normal Caller
      if is_header:
         is_header = False
      elif row[1].value is not None:
         # print(f'{row[0].value} -> {row[1].value}')
         # a list containing the guest name and note to the caller's list:
         caller_note = ""
         if row[3].value is not None and isinstance(row[3].value, str):
            caller_note = row[3].value.replace("’", "'")
         mapping_dict[row[1].value].append({'guest':row[0].value, 'caller_note':caller_note, 'change':row[2].value, 'normal_caller':row[4].value})
   # print(f"{mapping_dict=}\n")
   '''
   mapping_dict={'Caroline': [['Guest1', 'new regular']], 'Tina': [], 'Peter': [], 'Rebecca': [['Guest2', 'Substitute this week only']], 'Maria': [], 'Barb': [], 'Lisa': [['Guest3', None]], 'Do-Not-Call': []}
   '''
   callers_with_no_guest_list = []
   for caller, guests in mapping_dict.items():
      if len(guests) == 0:
         callers_with_no_guest_list.append(caller)
    # remove caller with no guests from mapping_dict
   for caller in callers_with_no_guest_list:   
         mapping_dict.pop(caller) 

   # make a dictionary of guest data to be used for generating reports.
   guest_dict = {}
   is_header = True
   invalid_usernames = []
   for row in workbook[GUESTS_SHEET_NAME].rows:
      if is_header:
         is_header = False
      elif row[GUEST_USERNAME].value is None and row[GUEST_LASTNAME].value is None:
         continue #skip blank rows
      else:
         # remove characters above the font range, e.g. "’", "“"
         cleaned_values = [""] * (GUEST_NOTES+1)
         for index in [GUEST_USERNAME, GUEST_FIRSTNAME, GUEST_LASTNAME, GUEST_PASSWORD, GUEST_TOWN, GUEST_PHONE, GUEST_NOTES]:
            if row[index].value is not None and isinstance(row[index].value, str):
               cleaned_values[index] = row[index].value.replace("’", "'")
               cleaned_values[index] = cleaned_values[index].replace("“",'"')
               cleaned_values[index] = cleaned_values[index].replace("”",'"')
               cleaned_values[index] = cleaned_values[index].replace("…",'"')
         if len(cleaned_values[GUEST_USERNAME]) < 3:
            invalid_usernames.append(f"{cleaned_values[GUEST_FIRSTNAME]} {cleaned_values[GUEST_LASTNAME]}")
         else:
            guest_dict[cleaned_values[GUEST_USERNAME]]= {'First':cleaned_values[GUEST_FIRSTNAME], 'Last':cleaned_values[GUEST_LASTNAME], 
            'PW':cleaned_values[GUEST_PASSWORD],'Town':cleaned_values[GUEST_TOWN], 'Phone':cleaned_values[GUEST_PHONE], 
            'Notes':cleaned_values[GUEST_NOTES]}      
   # print(f"{guest_dict=}")
   '''
   guest_dict={'Guest1': {'First': 'Guest', 'Last': 1.0, 'PW': 'secret', 'Town': 'Newbury', 'Phone': '978.555.0000', 'Notes': 'call early'}, 'Guest2': {'First': 'Guest', 'Last': 2.0, 'PW': 'secret', 'Town': 'Newbury', 'Phone': '978.555.0000', 'Notes': 'call 3 times'}, 'Guest3': {'First': 'Guest', 'Last': 3.0, 'PW': 'secret', 'Town': 'Newbury', 'Phone': '978.555.0000', 'Notes': 'call late'}}
   '''
   
   Caller_lists.caller_mapping_dict = mapping_dict
   Caller_lists.guest_dict = guest_dict
   Caller_lists.no_guest_list = callers_with_no_guest_list
   Caller_lists.invalid_usernames = invalid_usernames
   Caller_lists.success = True

   return Caller_lists


def filter_callers(caller_mapping_dict):
   filtered_dict = {}
   # see if the 'Change' row has a value, or if the caller does not equal the normal_caller
   callers_with_substitutes = []
   for caller, guests in caller_mapping_dict.items():
      changed = False
      for caller_guest_dict in guests:
         if caller_guest_dict['change'] is not None:
            changed = True
            # print(f"{caller}-{caller_guest_dict['guest']} {caller_guest_dict['change']=}")
         if caller != caller_guest_dict['normal_caller']:
            changed = True
            callers_with_substitutes.append(caller_guest_dict['normal_caller'])
            # print(f"{caller}-{caller_guest_dict['guest']} {caller_guest_dict['normal_caller']=}")
      if changed:
         filtered_dict[caller] = guests
         # print(f"filtered_dict_length={len(filtered_dict)}")

   # add callers_with_substitutes:
   for caller in callers_with_substitutes:
      filtered_dict[caller] = caller_mapping_dict[caller]
      
   return filtered_dict


def make_caller_pdfs(caller_mapping_dict, guest_dict, date_str, out_pdf_dir='.'):
   # PDF writing examples:
   #  https://medium.com/@mahijain9211/creating-a-python-class-for-generating-pdf-tables-from-a-pandas-dataframe-using-fpdf2-c0eb4b88355c
   #  https://py-pdf.github.io/fpdf2/Tutorial.html
   success_list = []
   failure_list = []
   print_count = 0
   for caller, guests in caller_mapping_dict.items():
      # print(f'{caller=} {guests=}')
      try:
         pdf = FPDF(orientation="L", format="letter") # default units are mm; adding , unit="in" inserts blank pages
         pdf.add_page()
         pdf.set_font("Helvetica", size=14)
         pdf.cell(0, 10, f"{caller} - Friday, {date_str}", align="C")
         pdf.ln(10)

         with pdf.table(col_widths=(11,13,14,13,13,14,14,30), line_height=6) as table:
            pdf.set_font(size=12)
            header = ['First', 'Last', 'UserName', 'Password', 'Town', 'Phone', 'Caller notes', 'Notes about guest']
            row = table.row()
            for column in header:
               row.cell(column)

            for caller_guest_dict in guests:
               if print_count > 0:
                  print(f'{caller_guest_dict=}')
                  print_count -= 1
               this_guest_username = caller_guest_dict['guest']
               if this_guest_username in guest_dict:
                  this_guest_dict = guest_dict[this_guest_username]
               else:
                  # print(f'{this_guest_username=} is not in guest_dict')
                  continue
               this_weeks_guest_note = caller_guest_dict['caller_note']
               if this_weeks_guest_note is None:
                  this_weeks_guest_note = ''
               row_data = [this_guest_dict['First'], this_guest_dict['Last'], this_guest_username, \
                           this_guest_dict['PW'], this_guest_dict['Town'], this_guest_dict['Phone'], \
                           this_weeks_guest_note, this_guest_dict['Notes']]
               # print(f"{row_data=}")
               row = table.row()
               for item in row_data:
                  row.cell(str(item), v_align="T")

         pdf.output(os.path.join(out_pdf_dir, f"{date_str}_{caller}.pdf"))
         success_list.append(caller)
      except Exception as e:
         try:
            current_app.logger.warning(f"PDF for {caller} failed: {e}")
         except:
            print(f"PDF for {caller} failed: {e}")
         failure_list.append(caller)
         # sys.exit(1)  # stop on first failure
   return (success_list, failure_list)


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


if __name__ == "__main__":
   argParser = argparse.ArgumentParser()
   argParser.add_argument("input", type=str, help="input filename with path")

   args = argParser.parse_args()

   if args.input is None:
      sys.exit("No file selected to parse.")
   elif Path(args.input).is_file():
      filename = args.input
   else:
      sys.exit("file selected is not a file.")

   Caller_lists = make_guests_per_caller_lists(filename)
   if not Caller_lists.success:
      print(f"Failure: {Caller_lists.message}")
      sys.exit(1)

   print(f"Callers with no guests: {Caller_lists.no_guest_list}")

   date_str = get_fridays_date_string()
   filtered_callers_dict = filter_callers(Caller_lists.caller_mapping_dict)
   make_caller_pdfs(filtered_callers_dict, Caller_lists.guest_dict, date_str, out_pdf_dir="/tmp")
   