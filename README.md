# A single web page asks for a excel file upload.

The upload_file() function in routes.py does all the work:
- saves the file
- runs a script to process the file which generates a new file
- downloads the resulting new file
- returns to the index page

_Launch using: flask run --host=0.0.0.0 --debugger_
