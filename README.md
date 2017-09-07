# docxtoctoexcel

Small program to convert the table of content of a docx file to an excel file. 
toexcel.match_apply contains the specific (adaptable) rules that define how the excel row has to look based on the toc-line structure.

## Usage:

Uses the following packages:

- openpyxl: https://openpyxl.readthedocs.io/en/default/
- python-docx: https://python-docx.readthedocs.io/en/latest/user/install.html
- tkinter: https://docs.python.org/3.6/library/tkinter.html

Main structure:

- fromdocx.toc_parser creates the rules to parse through the specified docx' file toc.
- toexcel.math_apply creates rules for the lay-out and content inserted into the xlsx file,
based on the content of the parsed toc.

Interface:

Uses a GUI to chose the docx file.