#Creating an index
####Objective: 
This program will scan a directory and all subdirectories looking for file names that contain the string set in the RecursiveFolder sub. In this search there are only excel files that meet the criteria so there is no requirement for file type filtering. When a file is found that matches the filter its path is added to the collection and the formula continues until all files are checked. 
Looping through the collection of paths, excel opens each workbook and loops through each sheet. There is a case statement for each file name which selects the columns from the different template files. The column data are then added to the current worksheet creating an index of all the data in the different worksheets.

####How to run:
- Open excel
- press ALT + F11 
- Create a new module 
- Copy this code
- Add a reference to Microsoft Scripting Runtime (Tools > References > Microsoft scripting runtime)
- Change the search directory in the code to the parent folder of where you want to search.
- Change the recursive filter to an appropriate filter for your files
- Run

####Requirements:
*This vba program is designed to run in excel and it is assumed the user has excel*
