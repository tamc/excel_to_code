# Cells

Within a worksheet xml, each cell is inidcated by a <c></c> tag:
  <c r="A1" t="b">
  	<f>TRUE</f>
  	<v>1</v>
  </c>

Where:

* r is the reference to that cell
* t is the optional type of the value of that cell
* <f></f> is optional and will contain the formula for that cell
* <v></v> is always present and will contain the value of that cell

## Value types

Value types are given by the 't' attibute of the <c> tag. 

If it is not present, then the cell's value is a number. No distinction is made between floats and integers.
	
Otherwise the attribute indicates:

* b - boolean, with the value either 1 for true or 0 for false
* s - shared string, with the value a number, starting at 0, giving the index of the string in the shared strings xml file
* str - normal string, with the value being that string
* e - error. The value will be the error that you see in excel as a string (.e.g., #NAME?)

### Format of value types in output

Produces value types in the format of referance\ttype\tvalue

If the cell's type is a number, then the type will contain n. Otherwise it will contain type indicators based on the above.

rewrite_values_to_include_shared_strings.rb can be used to turn value type 's' into value type 'str'