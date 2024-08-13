Excel Data Extraction Tool
Author: Georgiy Savchenko

Usage
Traverse the application using the tabs.

When you have filled out all fields, click "Add to table" to save the entry to the table
If you want to edit the entry, click the entry in the table, make your changes and then click "Update entry"
If you want to delete an entry, click the entry in the table, then click "Delete Entry"
When you have made all of the changes you want, click "Save to file" to save your changes permanently
Now you can click "Pull data" to perform your data Extraction

Extracting columns:
 1. Decide on a name, must be unique to the table
 2. Enabled: Decide whether this entry should be included in the Extraction
 3. Input filepath: the filepath to the excel workbook where the data should be extracted from
 4. Input sheetname: the sheetname in the excel workbook where the data should be extracted from
 5. Row offset: the minimum row the data should be extracted from. Rows lower than the row offset will not be included
 6. Columns: The columns that should be extracted, denoted by their letter in the excel spreadsheet
 7. Range Mode:
	- Dictates the range of rows that should be extracted
	8. End of Column: cells between the row offset and the end of column will be extracted
	9. Up to row: cells between the row offset and the row provided will be extracted
	10. Up to code: a code is defined in a column and the program will extract cells between the row offset and the row the code lies
 11. Output filepath: the filepath to the excel workbook where the data should be imported to
 12. Output sheetname: the sheetname in the excel workbook where the data should be imported to
 13. Mode:
	- Where in the output file the extracted data should be insert
	14. Next empty column: columns will be added at the end of the sheet
	15. Next empty row: columns will be added at the bottom of the sheet
 16. Add name: For each row that is exported, should the name of the entry appear of the left of the row

Extracting rows:
 1. Decide on a name, must be unique to the table
 2. Enabled: Decide whether this entry should be included in the Extraction
 3. Input filepath: the filepath to the excel workbook where the data should be extracted from
 4. Input sheetname: the sheetname in the excel workbook where the data should be extracted from
 5. Selection mode:
	- Defines which row should be extracted
	6. To row: extract a row by its index
	7. To code: extract the row which contains the code provided
 8. Output filepath: the filepath to the excel workbook where the data should be imported to
 9. Output sheetname: the sheetname in the excel workbook where the data should be imported to
 10. Add name: For each row that is exported, should the name of the entry appear of the left of the row

Parsing file:
 1. Decide on a name, must be unique to the table
 2. Enabled: Decide whether this entry should be included in the Extraction
 3. Input filepath: the filepath to the excel workbook where the data should be extracted from
 4. Input sheetname: the sheetname in the excel workbook where the data should be extracted from
 5. Column: parse the rows based on the values in this column
 8. Output filepath: the filepath to the excel workbook where the data should be imported to