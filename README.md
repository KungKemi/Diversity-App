# DiversityApp
1.1 OVERVIEW

The NCCD app was built with the purpose of compiling student data, from one or more Excel spreadsheets located in the same folder, into a
single Excel sheet.

1.2 RAW EXCEL SHEET LOCATION

It is expected that this folder is at a deeper branch of the current app directory (i.e., located in a folder in the same directory
as the app, or in a folder within said folder, etc.) For every additional level of branching, a \\\ must be appended after the name of 
the previous folder.

1.3 COLUMN FILE LOCATION

It is expected that a .txt file will be located in the same directory as the app, containing the names of any columns which you wish to extract.
Each column must be listed on a newline, with the column against which you want to group and sort the data on the first line (e.g., Student Name).

1.4 SCORING COLUMNS

Any columns containing numerical data are assumed to be scoring criteria, and hence will be used when discerning the student's mean category.
If a cell in a numerical column is empty, then it will be filled with a 0.

It is assumed that a column which contains numerical data will have one of the following values:

0 = None,
1 = Differentiation,
2 = Supplementary,
3 = Substantial,
4 = Extensive

If a column contains numerical data other than one of the five values specified above, then an error will be thrown.

If a numerical column contains non-numerical data, then an error will be thrown.

In the case that either of the two aforementioned errors are thrown, the problem cells will be reported and the user will be prompted to choose
a new folder. 

1.5 NAME COLUMN

The student data is expected to be grouped and sorted based on the student name. As such, any spelling errors could result in student data not being
properly compiled. Only the first and last names of the student are extracted from the names given, with any non-alphanumeric characters being discarded.
It is expected that names will be written in a specific format; i.e., "First Last" or "Last First".

Any rows which do not contain student names, but do contain numerical data or general comments, will be discarded.

1.6 MISSPELT NAMES

1.6.1 BUILT-IN SEARCH TOOL

In the event where student names are misspelt, the app can be used to optionally correct these misspellings in the created Excel sheet.
A similarity threshold is used to detect similar names, which is the minimum required percentage of similarity between two names before they are flagged.
This percentage is based off of the number and type of letters contained by both names. The lower the similarity threshold, the more names will be flagged.
Identical names will not be flagged.

It is recommended that a threshold of 85% is used as this will pick up most errors resulting from 1-3 spelling mistakes, or when the order of
first and last names vary.

When the app is asked to look for misspelt names, and two similar names are flagged, the user will be asked to pick one of three options:

1) Swap 'name one' with 'name two'
2) Swap 'name two' with 'name one'
3) Do nothing

In the case where option 1) is chosen, then all instances of 'name one' are replaced with 'name two'. The opposite is done if option 2) is chosen.
In the case where option 3) is chosen, all 'name one' and 'name two' pairs will remain unchanged and ignored thereafter.

1.6.2 SHEET COLUMN

Another means for identifying misspelt names can be through use of the 'Sheet' column in the resulting Excel sheet.
This column specifies the sheet from which each row of data is taken, and so any identified errors can be remedied in the offending sheet before
re-running the tool.

In order for this additional column to be displayed in the resulting sheet, 'Sheet' needs to be included on one of the lines of the text file containing
the column names.

1.7 OUTPUT

Upon running the program, an Excel sheet containing the processed data of the folder will be created in the same directory as the app.
It will be given the name 'Processed Data - folder name.xlsx'.

Within this file will be two sheets:

1) a 'combined' sheet containing all of the student data, ordered alphabetically and containing all of the data from the specified columns (if present);
   and
2) an 'average' sheet containing all of the student data grouped by name, with each numerical column being averaged and rounded to the nearest integer.
   A 'Mean' column will be included which specifies the category of the student after averaging the data across every numerical column for said student.


2.0 DISCLAIMER

Failure to use this app in the intended way, as stipulated above, may result in unexpected behaviour.
