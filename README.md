# LibrePyListEditor
Use Libre/Open Office Calc as Tree/Grid View for Browsing and Editing Python Nested List/Dict/Tuple Using UNO or COM 

This project is created for using Libre/Open office Calc Spread Sheet program as a grid view / tree view control for displaying and editing list, dictionary or tuple variables from Python command prompt. It provides a GUI tool for people who need to deal with long and deeply nested lists interactively on an Python commnad prompt.

This program can use either COM or UNO as the bridge to the Calc program, so it is not passing data thru files. 

What this program do is to display list variable in a simplified and tokenized Python syntax on spread sheet where each token is stored in a cell (it supports str, float and int). Nested list is shown using indentation and sub list can be expanded or collapsed interactively. The modified list can be read back from the Python command prompt after you finished changing it on the spread sheet.

## Dependencies
- Libre / Open Office (tested with Libre Office 5.4.2.2 (x64) on Windows 8.1, other versions and OSes should work) and the Python interpreter bundled in the Office installation. 
- A Python installation other than the one bundled with Office. It requires "comtypes" package to be installed (I used Anaconda Python installer which included this package). (Not necessary if you can use the bundled Python for your work.)

## A Simple Demo
1. Start Calc with special command line parameters to allow for control thru COM or UNO from Python interpreter.
   - On Windows, 
     - "C:\Program Files\LibreOffice 5\program\soffice.exe" "--calc" --accept="socket,host=localhost,port=2002;urp;"
   - On Linux, 

## References
