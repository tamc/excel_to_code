# Workbook.xml

This has two sections of interest:

* sheets - lists the names of the worksheets in the workbook
* definedNames - lists the named references in the workbook

## Format

	<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<fileVersion appName="xl" lastEdited="5" lowestEdited="4" rupBuild="20910"/>
		<workbookPr date1904="1" showInkAnnotation="0" autoCompressPictures="0"/>
		<bookViews>
			<workbookView xWindow="140" yWindow="0" windowWidth="33360" windowHeight="20300" tabRatio="500"/>
		</bookViews>
		<sheets>
			<sheet name="Outputs" sheetId="1" r:id="rId1"/>
			<sheet name="Calcs" sheetId="2" r:id="rId2"/>
			<sheet name="Inputs" sheetId="3" r:id="rId3"/>
		</sheets>
		<definedNames>
			<definedName name="In_result">Inputs!$A$3</definedName>
			<definedName name="Local_named_reference" localSheetId="3">Inputs!$A$3</definedName>
		</definedNames>
		<calcPr calcId="140001" calcMode="manual" concurrentCalc="0"/>
		<extLst>
			<ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="{7523E5D3-25F3-A5E0-1632-64F254C22452}">
				<mx:ArchID Flags="2"/>
			</ext>
		</extLst>
	</workbook>

## Notes

### Worksheet names

The r:id element refers to the associated relationships xml, which translates that r:id into a filename pointing at the worksheet xml


### Named references

These may have a 'localSheetId' attribute, whose value maps to the sheetId attribute of the sheet cells. In these cases, the named reference only applies within that worksheet. 

Because of the existence of the localSheetId, names need not be unique.



## Output

### Worksheet names

There are three steps to the extraction:

1. extract_worksheet_names.rb takes the workbook.xml and produces one worksheet per line. Relationship id then a tab then the worksheet name.
2. extract_relationships.rb operating on workbook.xml.rels produces one relationship per line. Relationship id then a tab then the filename.
3. rewrite_worksheet_names.rb operates on the results of 1 and 2 to produce one worksheet per line. Worksheet name, then a tab, then its filename.
