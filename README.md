# Room-Rental-PDF-Generator
It is just a simple rental contract pdf generator for my own usage.

# Dependency: 
mailmerge and comtypes. You can just install this through pip.

# Running: 
Open terminal or cmd and run `python generate_pdf.py`

# Note: 
If you want to make it into a excutable, you can just use cx_Freeze. I have include the setup.py in it, you can just run it by: `python setup.py build` in terminal or cmd.

# Feature: 
It produce a simple form that allows the user to input data and generate pdf file from the given .docx file/inputted data. I didn't include the doc file here, but you can feel free to make your own.

# Requirement for The Docx File:
- This program will try to update every field (it has to be merge field) in the doc. If you don't know how to create fields in word, check [this](https://www.webucator.com/how-to/how-insert-built-fields-microsoft-word.cfm). Also, whenever you add a new merge field and want to update it by using this program, you will need to change the code as well. I may add an addition feature that allows you to add new merge field through the form.

# Todo
- allow user to use gui to add new text entry.
