# Relationships

Relationships files map a relationship id to a filename. They are used to specify which files contain the xml for different sheets that are named in workbook.xml. They are also used to specify which files contain the xml for structured tables that are named in each worksheet.

There is a _rels directory at the top level that contains workbook.xml.rels which has the workbook to worksheet relationships.

There is a _rels directory within the worksheets folder that contains sheet1.xml.rels and similar, which contain the worksheet to table relationships.

Not every worksheet has a rels file.

## Format

Format of workbook.xml.rels:

    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    	<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
    	<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    	<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    	<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
    	<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
    	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    	<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
    </Relationships>

Format of a worksheet.xml.rels:

    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
    </Relationships>

## Notes

The relationship id always appears to be rId plus an integer. They don't seem to be reliably stored in ascending order.

The Type seems to be redundant at the moment, since the type is always inferable from the Target name.

The Target appears to be relative to the file that owns the relationships file (e.g., to workbook.xml or sheet1.xml, not to workbook.xml.rels or sheet1.xml.rels)

Unknown: can relationship id's be skipped?

## Output

Each relationsnip on a line. Each line has the relationship id, then a tab, then the target.
