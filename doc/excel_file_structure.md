# The structure of excel files

These documents describe .xlsx files written by Excel 2007 and later. It does not describe the older binary format.

An excel file is a zip collection of xml files.

Unzip the .xlsx file with:

    unzip -uo spreadsheet.xlsx -d directory_for_unzipped_files

Withn the top level of this folder will be:

* xl - The key folder, with details of all the worksheets
* _rels - Always empty?
* docProps - Not much of interest
* [Content Types].xml - Not sure what this is for

## xl 

Contains:

* _rels - Wihin this, workbook.xml.rels contians the mapping from sheet names (given in workbook.xml) to filenames
* calcChain.xml - Not relevant
* sharedStrings.xml - Strings used in worksheets are stored here, and then referenced from the worksheet by number
* styles.xml - Not relevant
* tables - Contians descriptions of the structured tables, which are referenced from individual worksheets
* theme - Not relevant
* workbook.xml - Contains the names of worksheets and named references
* worksheets - Contains the xml files describing individual worksheets

### xl / worksheets

There appear to be three ways that worksheets are named:

* A human readable name (e.g., "Costs sheet"), that matches the one seen in excel and excel references. This is listed in workbook.xml
* An index number (e.g., 3), given in workbook.xml
* A filename (e.g., sheet1.xml), given in _rels/workbook.xml.rels, that actually contains the formulae

The index number and filename may change if a worksheet is added or deleted. Not sure if it changes if worksheets are reordered.

