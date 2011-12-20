# SharedStrings.xml

Excel appears to store all strings that appear as cell values in a spreadsheet in a 'shared strings' document.

The exact filename is referenced in _rels/workbook.xml.rels, but always appears to be SharedStrings.xml in the top of the excel file.

## Format

Each string is wrapped in an `<si></si>` tag. The string may contain formatting.

    <?xml version="1.0"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
    	<si>
    		<t>This a second shared string</t>
    		<phoneticPr fontId="3" type="noConversion"/>
    	</si>
    	<si>
    		<r>
    			<t>This is, h</t>
    		</r>
    		<r>
    			<rPr>
    				<b/>
    				<sz val="10"/>
    				<rFont val="Verdana"/>
    			</rPr>
    			<t>opefully, the first</t>
    		</r>
    		<r>
    			<rPr>
    				<sz val="10"/>
    				<rFont val="Verdana"/>
    			</rPr>
    			<t xml:space="preserve"> shared string</t>
    		</r>
    		<phoneticPr fontId="3" type="noConversion"/>
    	</si>
    </sst>



## Notes on shared string referencing

When the type of a cell is 's' then the value of a cell will be an integer. That integer is the index of the shared string in this file, starting at zero.


## Output

The `extract_shared_strings.rb` script converts the shared strings xml into a text file. Each string appears on a separate line. Formatting tags and newlines are removed.

## Questions

Should we be preserving newlines in these strings?