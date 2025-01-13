# A single web page asks for a excel file upload.

The upload_file() function in routes.py does all the work:
- saves the file
- runs a script to process the file which generates a new file
- downloads the resulting new file
- returns to the index page

_Launch using: flask run --host=0.0.0.0 --debugger_

## Font problems
The text in the excel file had characters not supported by Helvetica, like curved single and double quotes.

Rather than do a string replace, then fonts could be loaded.  Refer to: 

https://py-pdf.github.io/fpdf2/Unicode.html

## references for going back to / after the file download
https://stackoverflow.com/questions/41518040/how-to-make-flask-to-send-a-file-and-then-redirect#47866877

