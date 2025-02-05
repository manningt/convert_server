# A single web page asks for a excel file upload.

The upload_file() function in routes.py does all the work:
- saves the excel file to /tmp/uploads
   - it expects the following nameing convention: Master List for Calling Feb 7 2025.xlsx
   - It extracts the string between Calling and .xlsx to use as the date for the files.
- runs a script to process the file and generates PDFs per caller and a processing report
- creates a zip file with the PDFs and report
- downloads the resulting zip file
- returns to the index page

_Launch using: flask run --host=0.0.0.0 --debugger_  >> after doing ```source .venv/bin/activate```

## a virtualenv needs to be created:
```
python3 -m venv .venv
pip install -r requirements.txt
```
## Font problems
The text in the excel file had characters not supported by Helvetica, like curved single and double quotes.

Rather than do a string replace, then fonts could be loaded.  Refer to: 

https://py-pdf.github.io/fpdf2/Unicode.html

## references for going back to / after the file download
https://stackoverflow.com/questions/41518040/how-to-make-flask-to-send-a-file-and-then-redirect#47866877
