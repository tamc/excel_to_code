# Cells

Within a worksheet xml, each cell is inidcated by a <c></c> tag:
  <c r="A1" t="b">
  	<f>TRUE</f>
  	<v>1</v>
  </c>

Where:

* r is the reference to that cell
* t is the optional type of the value of that cell
* <f></f> is optional and will contain the formula for that cell, with attributes indicating the formula type
* <v></v> is always present and will contain the value of that cell

## Value types

Value types are given by the optional 't' attibute of the <c> tag:

* No t attribute - the cell's value is a number. No distinction is made between floats and integers.
* b - boolean, with the value either 1 for true or 0 for false
* s - shared string, with the value a number, starting at 0, giving the index of the string in the shared strings xml file
* str - normal string, with the value being that string
* e - error. The value will be the error that you see in excel as a string (.e.g., #NAME?)

Note that the value of a cell is given in a <v></v> tag. If the <v> tag is self closing (e.g., <v/>) then the cell is blank.

### Format of value types in output

Produces value types in the format of referance\ttype\tvalue

If the cell's type is a number, then the type will contain n. Otherwise it will contain type indicators based on the above.

rewrite_values_to_include_shared_strings.rb can be used to turn value type 's' into value type 'str'

# Formula types

If present, the <f></f> tag will indicate a formula. It may have a 't' attribute to indicate the type of formula. Depending on the type of the formula it may have other attributes. The content is, in most cases, the formula as the user has entered it.

Note that the <f> tag may be self closing (<f/>). This means it doesn't specify a formula of its own. This is possible for shared and array formulae.

* No attribute - a straightforward formula
* shared - the same formula is shared amongst multiple cells
* array - an array formula

## shared formulae

A formula of type shared will always have an si attribute. This will contain an integer. All formula in a worksheet of type shared and that have the same integer will have the same formula, although the references will need to be adjusted for their relative position on the worksheet.

If a formula has a ref attribute, then it will contain the actual formula, an the ref attribute will indicate the range of cells that use it.

The ref attribute can refer to a single cell (e.g., A1) implying that the formula is not really shared.

If a formula doesn't have a ref attribute, then it will not contain any formula, and will expect to use one provided by a formula with both a ref attribute and the same si attribute.

It is a bit unclear why there is both an si attribute and a ref attribute.

The shared formula always appears to be in the top left of the ref range. Is this guaranteed?

## array formulae

An array formula will always have a ref attribute indicating the range of cells over which its answer will be spread. That range can be a single cell.

### Differences between formula as saved and formula as seen in excel

* References to external workbooks are modified