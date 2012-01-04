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

