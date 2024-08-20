# Excel Data Extraction Tool

A tool used to create profile for automatic data extraction and saving using excel files. Useful for those who consistently have to copy and paste rows, columns, etc from one excel file to another.

## Authors

- [Georgiy Savchenko](https://github.com/GeorgiySav)

## Chapters

- [Usage](#usage)
- [Field Information](#field-information)
- [Error Messages](#error-messages)
- [File Format](#file-format)
- [License](#license)

## Installation

Install from the "Releases" page. The zip file will contain a .exe file which is the application

If you would like to build from source, ensure you have the necessary dependencies:

- Python
- Pandas
- Openpyxl
- dearpygui
- xdialog

## Usage

To use the different pulling modes, traverse the application using the tabs at the top of the page. Each tabs' input fields and tables are independant of each other so don't worry about conflicts between each other.

### Modes

- Column
  - Extracts columns provided
- Row
  - Extracts a single row provided
- Parse
  - Splits rows of a table into seperate sheets based on a column provided

### Important notes

- Each entry contains a name and enabled attribute
- The name attribute must be unique compared to every other attribute in the table
- The enabled attribute determines whether the program should include this entry when going over every entry
- EVERYTHING IS CASE AND SPACE SENSITIVE

### How to add an entry

1. Fill out all of the fields.
2. Click "Add to table" to save the entry to the table.

### How to edit an entry

1. Click the entry you want to edit in the table.
2. Make your changes.
3. Click "Update entry" to save the new changes.

### How to delete an entry

1. Click the entry you want to delete from the table.
2. Click "Delete entry" to remove the entry from the table

### How to pull your data

1. Fill the table with the entries you want to pull
2. SAVE the entries to file by clicking "Save to file"
3. Click "Pull data"

## Field Information

### Column Entries

- Name
- Enabled
- Input filepath: the filepath to the excel workbook where the data should be extracted from
- Input sheetname: the sheetname in the excel workbook where the data should be extracted from
- Row offset: the minimum row index (lower bound) the data should be extracted from. Rows lower than the row offset will not be included
- Columns: The columns that should be extracted, denoted by their letter in the excel spreadsheet
- Range Mode: Dictates the range of cells that should be extracted
  - End of Column: cells between the row offset and the end of column will be extracted
  - Up to row: cells between the row offset and the row provided will be extracted
  - Up to code: a code is defined in a column and the program will extract cells between the row offset and the row the code lies
- Output filepath: the filepath to the excel workbook where the data should be imported to
- Output sheetname: the sheetname in the excel workbook where the data should be imported to
- Mode: Where in the output file the extracted data should be insert
  - Next empty column: columns will be added at the end of the sheet
  - Next empty row: columns will be added at the bottom of the sheet
- Add name: For each row that is exported, should the name of the entry appear of the left of the row

### Row Entries

- Name
- Enabled
- Input filepath: the filepath to the excel workbook where the data should be extracted from
- Input sheetname: the sheetname in the excel workbook where the data should be extracted from
- Selection mode: Defines which row should be extracted
  - To row: extract a row by its index
  - To code: extract the row which contains the code provided
- Output filepath: the filepath to the excel workbook where the data should be imported to
- Output sheetname: the sheetname in the excel workbook where the data should be imported to
- Add name: For each row that is exported, should the name of the entry appear of the left of the row

### Parse Entries

- Name
- Enabled
- Input filepath: the filepath to the excel workbook where the data should be extracted from
- Input sheetname: the sheetname in the excel workbook where the data should be extracted from
- Column: parse the rows based on the values in this column. Cells that share values are saved to a sheet with that name. Also excel, only allows sheet names to be up to 30 characters long, so the names will be cut from the beginning to fit
- Output filepath: the filepath to the excel workbook where the data should be imported to

## Field Information

### Column Entries

- Name
- Enabled
- Input filepath: the filepath to the excel workbook where the data should be extracted from
- Input sheetname: the sheetname in the excel workbook where the data should be extracted from
- Row offset: the minimum row index (lower bound) the data should be extracted from. Rows lower than the row offset will not be included
- Columns: The columns that should be extracted, denoted by their letter in the excel spreadsheet
- Range Mode: Dictates the range of cells that should be extracted
  - End of Column: cells between the row offset and the end of column will be extracted
  - Up to row: cells between the row offset and the row provided will be extracted
  - Up to code: a code is defined in a column and the program will extract cells between the row offset and the row the code lies
- Output filepath: the filepath to the excel workbook where the data should be imported to
- Output sheetname: the sheetname in the excel workbook where the data should be imported to
- Mode: Where in the output file the extracted data should be insert
  - Next empty column: columns will be added at the end of the sheet
  - Next empty row: columns will be added at the bottom of the sheet
- Add name: For each row that is exported, should the name of the entry appear of the left of the row

### Row Entries

- Name
- Enabled
- Input filepath: the filepath to the excel workbook where the data should be extracted from
- Input sheetname: the sheetname in the excel workbook where the data should be extracted from
- Selection mode: Defines which row should be extracted
  - To row: extract a row by its index
  - To code: extract the row which contains the code provided
- Output filepath: the filepath to the excel workbook where the data should be imported to
- Output sheetname: the sheetname in the excel workbook where the data should be imported to
- Add name: For each row that is exported, should the name of the entry appear of the left of the row

### Parse Entries

- Name
- Enabled
- Input filepath: the filepath to the excel workbook where the data should be extracted from
- Input sheetname: the sheetname in the excel workbook where the data should be extracted from
- Column: parse the rows based on the values in this column. Cells that share values are saved to a sheet with that name. Also excel, only allows sheet names to be up to 30 characters long, so the names will be cut from the beginning to fit
- Output filepath: the filepath to the excel workbook where the data should be imported to

## Error Messages

At the bottom of the page, there is a status report window that reports the success and errors when you pull entries.

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

## File format

Entries are saved in a json file so that the user doesn't have to re-enter fields everytime they open the application

### Column format

```json
In project_info.json
[
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
    },
    ...
]
```

### Row format

```json
In row_project_info.json
[
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
    ...
]
```

### Parse format

```json
In parse_project_info.json
[
    {
        "name": "project name",
        "enabled": true,
        "input": {
        "filepath": "C:\\...",
        "sheetname": "Sheet1",
        "column": "ZZ"
        },
        "output": {
        "filepath": "C:\\..."
        }
    },
    ...
]
```

## License

[MIT](https://choosealicense.com/licenses/mit/)
