Excel Data Extraction Tool
Author: Georgiy Savchenko

Usage
Traverse the application using the tabs.
Each tabs input fields and tables are independant of each other

When you have filled out all fields, click "Add to table" to save the entry to the table. Check below on the explanations for each field
If you want to edit the entry, click the entry in the table, make your changes and then click "Update entry"
If you want to delete an entry, click the entry in the table, then click "Delete Entry"
When you have made all of the changes you want, click "Save to file" to save your changes permanently
Now you can click "Pull data" to perform your data Extraction
You will see the status of each entry in the section at the bottom. If anything goes wrong, it should explain what the issue is

!!!!!!!!!!!!!!!EVERYTHING IS CASE AND SPACE SENSITIVE!!!!!!!!!!!!!!!!!
!!!IT WILL BE EASIER TO COPY AND PASTE VALUES THAN TYPE IT YOURSELF!!!

Extracting columns
 1.  Name: must be unique to the table
 2.  Enabled: Decide whether this entry should be included in the extraction when you click the pull button
 3.  Input filepath: the filepath to the excel workbook where the data should be extracted from
 4.  Input sheetname: the sheetname in the excel workbook where the data should be extracted from
 5.  Row offset: the minimum row the data should be extracted from. Rows lower than the row offset will not be included
 6.  Columns: The columns that should be extracted, denoted by their letter in the excel spreadsheet
 7.  Range Mode:
	- Dictates the range of cells that should be extracted
	8.  End of Column: cells between the row offset and the end of column will be extracted
	9.  Up to row: cells between the row offset and the row provided will be extracted
	10. Up to code: a code is defined in a column and the program will extract cells between the row offset and the row the code lies
 11. Output filepath: the filepath to the excel workbook where the data should be imported to
 12. Output sheetname: the sheetname in the excel workbook where the data should be imported to
 13. Mode:
	- Where in the output file the extracted data should be insert
	14. Next empty column: columns will be added at the end of the sheet
	15. Next empty row: columns will be added at the bottom of the sheet
 16. Add name: For each row that is exported, should the name of the entry appear of the left of the row
File entries are save to: project_info.json
entry format:
{
	"name": "project name",
	"enabled": true or false,
	"input": {
		"filepath": "C://...",
		"sheetname": "sheet1",
		"row_offset": 1,
		"columns": [
			"A", "B", ...
		],
		"range_mode": {
        	"type": "End of column"
      	} or {
			"type": "Up to code",
        	"code": "stop",
        	"column": "C"
		} or {
			"type": "Up to row",
        	"row": 155
		}
	}
	"output": {
		"filepath": "C://...",
		"sheetname": "sheet1",
		"insert_mode": "Next empty row" or "Next empty column",
      	"include_name": true or false
	}
}

Extracting rows:
 1.  Decide on a name, must be unique to the table
 2.  Enabled: Decide whether this entry should be included in the Extraction
 3.  Input filepath: the filepath to the excel workbook where the data should be extracted from
 4.  Input sheetname: the sheetname in the excel workbook where the data should be extracted from
 5.  Selection mode:
	- Defines which row should be extracted
	6. To row: extract a row by its index
	7. To code: extract the row which contains the code provided
 8.  Output filepath: the filepath to the excel workbook where the data should be imported to
 9.  Output sheetname: the sheetname in the excel workbook where the data should be imported to
 10. Add name: For each row that is exported, should the name of the entry appear of the left of the row
File entries are save to: row_project_info.json
entry format:
{
	"name": "project name",
    "enabled": true,
    "input": {
    	"filepath": "C:\\...",
    	"sheetname": "sheet1",
    	"selection_mode": {
    		"type": "code",
        	"code": "Total",
        	"column": "B"
      	} or {
			"type": "row",
			"row": 1
		}
    },
    "output": {
    	"filepath": "C:\\...",
    	"sheetname": "sheet1",
    	"include_name": true
    }
},

Parsing file:
 1.  Decide on a name, must be unique to the table
 2.  Enabled: Decide whether this entry should be included in the Extraction
 3.  Input filepath: the filepath to the excel workbook where the data should be extracted from
 4.  Input sheetname: the sheetname in the excel workbook where the data should be extracted from
 5.  Column: parse the rows based on the values in this column. Cells that share values are saved to a sheet with that name
 8.  Output filepath: the filepath to the excel workbook where the data should be imported to
 File entries are save to: parse_project_info.json
 entry format:
 {
    "name": "Test",
    "enabled": true,
    "input": {
      "filepath": "C:\\Users\\GeorgiySavchenko\\Documents\\Pull From SS\\test_output.xlsx",
      "sheetname": "Sheet1",
      "column": "ZZ"
    },
    "output": {
      "filepath": "C:\\Users\\GeorgiySavchenko\\Documents\\Pull From SS\\test_input.xlsx"
    }
}

Error Messages:
At the bottom of the application, there is a status report window that reports the success and errors when you pull entries. 
Error Messages and their meanings:
 - Failed to open the parse files: ...
   - There's an issue with the files that store the entries.
   - The file name could've changed from what it should be
   - Or the formatting has been corrupted. In this scenario you can attempt to fix the formatting, however if you have no experience with dealing with JSON files, then it might be very difficult. Otherwise you should replace the entire contents of the file with a pair of square brackets: []
 - Failed to open the input file: 
   - There has been an issue with opening the input excel spreadsheet. The status report should be able to provide a good enough reason for why it wasn't able to open to file
   - Just make sure the filepath is correct and the sheetname is also exactly the same as the one in the file. Remember the values are case and space sensitive
 - Failed to open the output file:
   - The same as the previous error but with the output excel spreadsheet.
 - Columns provided aren't within the range in the input file:
   - At least one of the columns provided to be extracted is outside of the range of the excel sheet. 
   - Make sure that the column isn't empty
 - Row provided isn't within range:
   - Row index provided is outside of the range of the sheet
   - The row index cannot be higher than the last row
 - Offset row provided is out of range:
   - Same as the previous error
 - Failed to write to the output file: ...
   - There's an issue with saving the changes to the output file.
   - Make sure that the output file is closed