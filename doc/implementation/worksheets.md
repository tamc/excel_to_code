# Worksheets

Contains the following elements:

1. dimension - indicates the range of cells that are active in this worksheet (i.e., they have data)
2. sheetViews - indicates the currently visible area of the worksheet
3. sheetFormatPr - ?
4. sheetData - contains descriptions of each cell in the worksheet. See the cell.md document for more info.
5. pageMargins - self explanatory
6. tableParts - references to the structured tables that are defined on this workshet (see table documentation)
7. extLst - ?

## Outputs

extract/extract_worksheet_dimensions.rb will extract the dimensions of a worksheet
rewrite/rewrite_whole_row_column_references_to_areas.rb will use those dimensions to convert references of the form C:E and 12:15 into those of the form C1:E80 and A12:AA15
