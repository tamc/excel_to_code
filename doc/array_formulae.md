# Array formulae

As far as I can tell, there are two sets of rules that relate to array formulas: how they are presented and how they are calculated

## Presentation of array formulae

This is the easy part: Array formulae can be spread over more than one cell. This allows them to access different elements of any arrays that a function returns.

In the rewrite stage, I've chosen to implement this in a similar fashion to google docs: 

1. The first occurence of the array formula gets wrapped in the function ARRAYFORMULA()
2. Subsequent occuerences of the array formula get given the function CONTINUE(<ref to array formula>,<row offset>,<column offset>)
	
## Calculation of array formulae

This is the tricky part, because I fundamentally don't get why SUM({1,2}+{1,2}) works anytime but SUM(A1:A2+B1:B2) only works if it is specified as being an array formula.

So, my planned implementation would be to ensure that all functions can take arrays as arguments. The default approach to taking a series of arrays as arguments is to:

1. Expand all arguments so that their arrays are of equivalent size
2. Createa a new result array of that size
3. Populate that result array by iterating for each point of the array, applying the function to the values of the argument arrays at the same point


## Notes on how array formulae work

The default is that all operators can take an array. Therefore A1:A2+B1:B2 maps to {A1+B1;A2+B2}, althogugh the later isn't legal in excel formulae

The same is true of functions that expect single values. Therefore COSH(A1:A2) maps to {COSH(A1);COSH(A2)}

Some functions have arguments that are expected to be arrays, therefore they are left alone: SUM(A1:A2) maps to SUM(A1:A2)

Some functions have mixture of arguments, some are expected to be single values, others ranges. In this case, a mixture of things happen:

  INDEX(A1:A2,B1:B2) maps to {INDEX(A1,B1:B2); INDEX(B1,B1:B2)}
  
The tricky bit occurs when functions are nested:

  SUM(A1:A2+B1:B2) maps to SUM({A1+B1;A2+B2}) # Ok
  SUM(COSH(A1:A2)) maps to SUM({COSH(A1);COSH(A2)}) # Ok
  SUM(INDEX(A1:A2,B1:B2)) maps to  {SUM(INDEX(A1,B1:B2)); SUM(INDEX(B1,B1:B2))} # wtf?
  SUM(COUNTIF(A1:A2,B1:B2)) maps to SUM({COUNTIF(A1:A2,B1);COUNTIF(A1:A2,B1)}) # Ok