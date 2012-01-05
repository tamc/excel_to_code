# Tables

Tables are defined in individual files in the `xl\tables` directory. 

To know what sheet a table relates to, you need to make use of the <tablePart> section of a worksheet:

	<tableParts count="3">
		<tablePart r:id="rId1"/>
		<tablePart r:id="rId2"/>
		<tablePart r:id="rId3"/>
	</tableParts>

This gives a series of relationship ids, that can be translated into filenames using the associated relationship file in `xl\worksheets\_rels`:

	<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table36.xml"/>
		<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table37.xml"/>
		<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table38.xml"/>
	</Relationships>

The table*.xml file then contains all information required to define the table, except its worksheet name:

	<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="16" name="Global.Assumptions" displayName="Global.Assumptions" ref="B5:G16" totalsRowCount="1" headerRowDxfId="6092" dataDxfId="6091">
		<autoFilter ref="B5:G15"/>
		<tableColumns count="6">
			<tableColumn id="1" name="Year" dataDxfId="6090" totalsRowDxfId="6089"/>
			<tableColumn id="2" name="Population" totalsRowFunction="average" dataDxfId="6088" totalsRowDxfId="6087" dataCellStyle="Comma"/>
			<tableColumn id="3" name="Households" dataDxfId="6086" totalsRowDxfId="6085" dataCellStyle="Comma">
				<calculatedColumnFormula>Global.Assumptions[[#This Row],[Population]]/25751000</calculatedColumnFormula>
			</tableColumn>
			<tableColumn id="4" name="GDP (2005 £m)" dataDxfId="6084" totalsRowDxfId="6083"/>
			<tableColumn id="5" name="GDP (£2010)" dataDxfId="6082" totalsRowDxfId="6081" dataCellStyle="Comma">
				<calculatedColumnFormula>Global.Assumptions[[#This Row],[GDP (2005 £m)]]*MGBP*Price2005</calculatedColumnFormula>
			</tableColumn>
			<tableColumn id="6" name="GDP per Capita (£2010)" dataDxfId="6080" totalsRowDxfId="6079">
				<calculatedColumnFormula>Global.Assumptions[[#This Row],[GDP (£2010)]]/Global.Assumptions[[#This Row],[Population]]</calculatedColumnFormula>
			</tableColumn>
		</tableColumns>
		<tableStyleInfo name="EnergyCalcTables" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/>
	</table>

In particular, within the <table> tag, the <displayName> attribute appears to be the name actually used in formulae, the <ref> attribute is the area covered by the table, including the header row and any total rows. <totalRowCount> is the number of rows at the bottom of the area that are taken up with totals rather than with data.
	
The <tableColumn> have a <name> attribute which is the name of the column as used in formulae. They appear to be listed in their left right order.

## Gotchas

Table names are not case sensitive
