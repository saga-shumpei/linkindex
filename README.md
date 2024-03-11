# linkindex
A macro that creates a link to any location on Excel and also creates a link on the log file.
## Preparing to use linkindex
1. Create a new book with excel.
2. Save it as "C:\Users\XXXX\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB". (Change XXXX to match your username.)
![macrobook](macrobook.png)
3. Download "linkindex.bas" to your PC.
4. Open Visual Basic and import "linkindex.bas".
![import](import.png)
6. Create an excel workbook for history and place it in any directory.
![logfilepath](logfilepath.png)
7. The history excel workbook must have a sheet named "latest".
![historyworkbook](historyworkbook.png)
8. A sheet called "settings" must exist in "PERSONAL.XLSB". A CELL with the name "LogFilePath" must exist in the "settings" sheet. Set the path of the excel workbook for history in the CELL with the name "LogFilePath".
![personal](personal.png)
9. Assign the NameLinkLog sub procedure of "linkindex.bas" to a shortcut key.
![shortcut](shortcut.png)
![shortcut2](shortcut2.png)
## How to use linkindex
10. The Excel workbook managed by linkindex must have a sheet named "index".
![managedbook](managedbook.png)
11. In the Excel workbook managed by linkindex, place the cursor on the CELL you want to add a heading to and press the shortcut key set in step 9.
![type](type.png)
12. The index sheet will then be displayed, so click anywhere and the index will be entered there.
![inputindex](inputindex.png)
![inputindex2](inputindex2.png)
13. The display returns to the CELL with heading in the Excel workbook managed by linkindex. At this time, this book has already been saved.
![return](return.png)
