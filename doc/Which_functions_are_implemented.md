# Which functions are implemented?

Below is a list of Excel Functions taken from: http://office.microsoft.com/en-us/excel-help/excel-functions-alphabetical-list-HA010277524.aspx plus some of the extra functions for 2007 and 2011 added manually.

Those that have been implemented have an _R_ (if they have been implemented in ruby) and a _C_ (if they have been implemented in c) after their name.

To see the implementation:
* ruby versions: src/excel/excel_functions and spec/excel/excel_functions
* C versions: src/compile/c/excel_to_c_runtime.c (tests are at the foot)

To add a missing function, read: doc/How_to_add_a_missing_function.md

## Database functions

* DAVERAGE -- Returns the average of selected database entries
* DCOUNT -- Counts the cells that contain numbers in a database
* DCOUNTA -- Counts nonblank cells in a database
* DGET -- Extracts from a database a single record that matches the specified criteria
* DMAX -- Returns the maximum value from selected database entries
* DMIN -- Returns the minimum value from selected database entries
* DPRODUCT -- Multiplies the values in a particular field of records that match the criteria in a database
* DSTDEV -- Estimates the standard deviation based on a sample of selected database entries
* DSTDEVP -- Calculates the standard deviation based on the entire population of selected database entries
* DSUM -- Adds the numbers in the field column of records in the database that match the criteria
* DVAR -- Estimates variance based on a sample from selected database entries
* DVARP -- Calculates variance based on the entire population of selected database entries

## Date and time functions

* DATE -- Returns the serial number of a particular date
* DATEVALUE -- Converts a date in the form of text to a serial number
* DAY -- Converts a serial number to a day of the month
* DAYS360 -- Calculates the number of days between two dates based on a 360-day year
* EDATE -- Returns the serial number of the date that is the indicated number of months before or after the start date
* EOMONTH -- Returns the serial number of the last day of the month before or after a specified number of months
* HOUR -- Converts a serial number to an hour
* MINUTE -- Converts a serial number to a minute
* MONTH -- Converts a serial number to a month
* NETWORKDAYS -- Returns the number of whole workdays between two dates
* NOW -- Returns the serial number of the current date and time
* SECOND -- Converts a serial number to a second
* TIME -- Returns the serial number of a particular time
* TIMEVALUE -- Converts a time in the form of text to a serial number
* TODAY -- Returns the serial number of today's date
* WEEKDAY -- Converts a serial number to a day of the week
* WEEKNUM -- Converts a serial number to a number representing where the week falls numerically with a year
* WORKDAY -- Returns the serial number of the date before or after a specified number of workdays
* YEAR -- Converts a serial number to a year
* YEARFRAC -- Returns the year fraction representing the number of whole days between start_date and end_date

## Engineering functions

* BESSELI -- Returns the modified Bessel function In(x)
* BESSELJ -- Returns the Bessel function Jn(x)
* BESSELK -- Returns the modified Bessel function Kn(x)
* BESSELY -- Returns the Bessel function Yn(x)
* BIN2DEC -- Converts a binary number to decimal
* BIN2HEX -- Converts a binary number to hexadecimal
* BIN2OCT -- Converts a binary number to octal
* COMPLEX -- Converts real and imaginary coefficients into a complex number
* CONVERT -- Converts a number from one measurement system to another
* DEC2BIN -- Converts a decimal number to binary
* DEC2HEX -- Converts a decimal number to hexadecimal
* DEC2OCT -- Converts a decimal number to octal
* DELTA -- Tests whether two values are equal
* ERF -- Returns the error function
* ERFC -- Returns the complementary error function
* GESTEP -- Tests whether a number is greater than a threshold value
* HEX2BIN -- Converts a hexadecimal number to binary
* HEX2DEC -- Converts a hexadecimal number to decimal
* HEX2OCT -- Converts a hexadecimal number to octal
* IMABS -- Returns the absolute value (modulus) of a complex number
* IMAGINARY -- Returns the imaginary coefficient of a complex number
* IMARGUMENT -- Returns the argument theta, an angle expressed in radians
* IMCONJUGATE -- Returns the complex conjugate of a complex number
* IMCOS -- Returns the cosine of a complex number
* IMDIV -- Returns the quotient of two complex numbers
* IMEXP -- Returns the exponential of a complex number
* IMLN -- Returns the natural logarithm of a complex number
* IMLOG10 -- Returns the base-10 logarithm of a complex number
* IMLOG2 -- Returns the base-2 logarithm of a complex number
* IMPOWER -- Returns a complex number raised to an integer power
* IMPRODUCT -- Returns the product of from 2 to 29 complex numbers
* IMREAL -- Returns the real coefficient of a complex number
* IMSIN -- Returns the sine of a complex number
* IMSQRT -- Returns the square root of a complex number
* IMSUB -- Returns the difference between two complex numbers
* IMSUM -- Returns the sum of complex numbers
* OCT2BIN -- Converts an octal number to binary
* OCT2DEC -- Converts an octal number to decimal
* OCT2HEX -- Converts an octal number to hexadecimal

## Financial functions

* ACCRINT -- Returns the accrued interest for a security that pays periodic interest
* ACCRINTM -- Returns the accrued interest for a security that pays interest at maturity
* AMORDEGRC -- Returns the depreciation for each accounting period by using a depreciation coefficient
* AMORLINC -- Returns the depreciation for each accounting period
* COUPDAYBS -- Returns the number of days from the beginning of the coupon period to the settlement date
* COUPDAYS -- Returns the number of days in the coupon period that contains the settlement date
* COUPDAYSNC -- Returns the number of days from the settlement date to the next coupon date
* COUPNCD -- Returns the next coupon date after the settlement date
* COUPNUM -- Returns the number of coupons payable between the settlement date and maturity date
* COUPPCD -- Returns the previous coupon date before the settlement date
* CUMIPMT -- Returns the cumulative interest paid between two periods
* CUMPRINC -- Returns the cumulative principal paid on a loan between two periods
* DB -- Returns the depreciation of an asset for a specified period by using the fixed-declining balance method
* DDB -- Returns the depreciation of an asset for a specified period by using the double-declining balance method or some other method that you specify
* DISC -- Returns the discount rate for a security
* DOLLARDE -- Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number
* DOLLARFR -- Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction
* DURATION -- Returns the annual duration of a security with periodic interest payments
* EFFECT -- Returns the effective annual interest rate
* FV -- Returns the future value of an investment
* FVSCHEDULE -- Returns the future value of an initial principal after applying a series of compound interest rates
* INTRATE -- Returns the interest rate for a fully invested security
* IPMT -- Returns the interest payment for an investment for a given period
* IRR -- Returns the internal rate of return for a series of cash flows
* ISPMT -- Calculates the interest paid during a specific period of an investment
* MDURATION -- Returns the Macauley modified duration for a security with an assumed par value of $100
* MIRR -- Returns the internal rate of return where positive and negative cash flows are financed at different rates
* NOMINAL -- Returns the annual nominal interest rate
* NPER -- Returns the number of periods for an investment
* NPV -- Returns the net present value of an investment based on a series of periodic cash flows and a discount rate
* ODDFPRICE -- Returns the price per $100 face value of a security with an odd first period
* ODDFYIELD -- Returns the yield of a security with an odd first period
* ODDLPRICE -- Returns the price per $100 face value of a security with an odd last period
* ODDLYIELD -- Returns the yield of a security with an odd last period
* PMT _R_ _C_ -- Returns the periodic payment for an annuity
* PPMT -- Returns the payment on the principal for an investment for a given period
* PRICE -- Returns the price per $100 face value of a security that pays periodic interest
* PRICEDISC -- Returns the price per $100 face value of a discounted security
* PRICEMAT -- Returns the price per $100 face value of a security that pays interest at maturity
* PV -- Returns the present value of an investment - _R_ _C_
* RATE -- Returns the interest rate per period of an annuity
* RECEIVED -- Returns the amount received at maturity for a fully invested security
* SLN -- Returns the straight-line depreciation of an asset for one period
* SYD -- Returns the sum-of-years' digits depreciation of an asset for a specified period
* TBILLEQ -- Returns the bond-equivalent yield for a Treasury bill
* TBILLPRICE -- Returns the price per $100 face value for a Treasury bill
* TBILLYIELD -- Returns the yield for a Treasury bill
* VDB -- Returns the depreciation of an asset for a specified or partial period by using a declining balance method
* XIRR -- Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic
* XNPV -- Returns the net present value for a schedule of cash flows that is not necessarily periodic
* YIELD -- Returns the yield on a security that pays periodic interest
* YIELDDISC -- Returns the annual yield for a discounted security; for example, a Treasury bill
* YIELDMAT -- Returns the annual yield of a security that pays interest at maturity

## Information functions

* CELL -- Returns information about the formatting, location, or contents of a cell - filename info_source implemented in ruby
* ERROR.TYPE -- Returns a number corresponding to an error type
* INFO -- Returns information about the current operating environment
* ISBLANK -- Returns TRUE if the value is blank
* ISERR -- Returns TRUE if the value is any error value except ##N/A
* ISERROR -- Returns TRUE if the value is any error value
* ISEVEN -- Returns TRUE if the number is even
* ISLOGICAL -- Returns TRUE if the value is a logical value
* ISNA -- Returns TRUE if the value is the ##N/A error value
* ISNONTEXT -- Returns TRUE if the value is not text
* ISNUMBER -- Returns TRUE if the value is a number
* ISODD -- Returns TRUE if the number is odd
* ISREF -- Returns TRUE if the value is a reference
* ISTEXT -- Returns TRUE if the value is text
* N -- Returns a value converted to a number
* NA -- Returns the error value ##N/A
* TYPE -- Returns a number indicating the data type of a value

## Logical functions

* AND -- Returns TRUE if all of its arguments are TRUE -- _R_ _C_
* FALSE -- Returns the logical value FALSE
* IF -- Specifies a logical test to perform -- _R_ _C_
* NOT -- Reverses the logic of its argument -- _R_ _C_ 
* OR -- Returns TRUE if any argument is TRUE
* TRUE -- Returns the logical value TRUE

## Lookup and reference functions

* ADDRESS -- Returns a reference as text to a single cell in a worksheet
* AREAS -- Returns the number of areas in a reference
* CHOOSE -- _R_ _C_ Chooses a value from a list of values
* COLUMN -- Returns the column number of a reference
* COLUMNS -- Returns the number of columns in a reference
* GETPIVOTDATA -- Returns data stored in a PivotTable
* HLOOKUP -- Looks in the top row of an array and returns the value of the indicated cell -- _R_ _C_
* HYPERLINK -- Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the Internet
* INDEX -- Uses an index to choose a value from a reference or array -- _R_ _C_ 
* INDIRECT -- Returns a reference indicated by a text value -- _R_ _C_ only implemented for indirects that can be converted to formula at compile time
* LOOKUP -- Looks up values in a vector or array
* MATCH -- Looks up values in a reference or array -- _R_ _C_
* OFFSET -- Returns a reference offset from a given reference
* ROW -- Returns the row number of a reference
* ROWS -- Returns the number of rows in a reference
* RTD -- Retrieves real-time data from a program that supports COM automation
* TRANSPOSE -- Returns the transpose of an array
* VLOOKUP -- Looks in the first column of an array and moves across the row to return the value of a cell -- _R_ _C_

## Math and trigonometry functions

* ABS -- Returns the absolute value of a number -- _R_ _C_
* ACOS -- Returns the arccosine of a number
* ACOSH -- Returns the inverse hyperbolic cosine of a number
* ASIN -- Returns the arcsine of a number
* ASINH -- Returns the inverse hyperbolic sine of a number
* ATAN -- Returns the arctangent of a number
* ATAN2 -- Returns the arctangent from x- and y-coordinates
* ATANH -- Returns the inverse hyperbolic tangent of a number
* CEILING -- Rounds a number to the nearest integer or to the nearest multiple of significance
* COMBIN -- Returns the number of combinations for a given number of objects
* COS -- Returns the cosine of a number
* COSH -- Returns the hyperbolic cosine of a number -- _R_ _C_
* DEGREES -- Converts radians to degrees
* EVEN -- Rounds a number up to the nearest even integer
* EXP -- Returns e raised to the power of a given number
* FACT -- Returns the factorial of a number
* FACTDOUBLE -- Returns the double factorial of a number
* FLOOR -- Rounds a number down, toward zero
* GCD -- Returns the greatest common divisor
* INT -- Rounds a number down to the nearest integer -- _R_ _C_
* LCM -- Returns the least common multiple
* LN -- Returns the natural logarithm of a number
* LOG -- Returns the logarithm of a number to a specified base
* LOG10 -- Returns the base-10 logarithm of a number
* MDETERM -- Returns the matrix determinant of an array
* MINVERSE -- Returns the matrix inverse of an array
* MMULT -- Returns the matrix product of two arrays -- _R_ _C_ but posibly only when as array formula
* MOD -- Returns the remainder from division -- _R_ _C_
* MROUND -- Returns a number rounded to the desired multiple
* MULTINOMIAL -- Returns the multinomial of a set of numbers
* ODD -- Rounds a number up to the nearest odd integer
* PI -- Returns the value of pi
* POWER -- Returns the result of a number raised to a power -- _R_ _C_
* PRODUCT -- Multiplies its arguments
* QUOTIENT -- Returns the integer portion of a division
* RADIANS -- Converts degrees to radians
* RAND -- Returns a random number between 0 and 1
* RANDBETWEEN -- Returns a random number between the numbers you specify
* ROMAN -- Converts an arabic numeral to roman, as text
* ROUND -- Rounds a number to a specified number of digits -- _R_ _C_
* ROUNDDOWN -- Rounds a number down, toward zero -- _R_ _C_
* ROUNDUP -- Rounds a number up, away from zero  -- _R_ _C_
* SERIESSUM -- Returns the sum of a power series based on the formula
* SIGN -- Returns the sign of a number
* SIN -- Returns the sine of the given angle
* SINH -- Returns the hyperbolic sine of a number
* SQRT -- Returns a positive square root
* SQRTPI -- Returns the square root of (number * pi)
* SUBTOTAL -- Returns a subtotal in a list or database  -- _R_ _C_
* SUM -- Adds its arguments -- _R_ _C_
* SUMIF -- Adds the cells specified by a given criteria -- _R_ _C_
* SUMIFS -- Adds the cells specified by a given criteria -- _R_ _C_
* SUMPRODUCT -- Returns the sum of the products of corresponding array components -- _R_ _C_
* SUMSQ -- Returns the sum of the squares of the arguments
* SUMX2MY2 -- Returns the sum of the difference of squares of corresponding values in two arrays
* SUMX2PY2 -- Returns the sum of the sum of squares of corresponding values in two arrays
* SUMXMY2 -- Returns the sum of squares of differences of corresponding values in two arrays
* TAN -- Returns the tangent of a number
* TANH -- Returns the hyperbolic tangent of a number
* TRUNC -- Truncates a number to an integer

## Statistical functions

* AVEDEV -- Returns the average of the absolute deviations of data points from their mean
* AVERAGE -- Returns the average of its arguments -- _R_ _C_
* AVERAGEA -- Returns the average of its arguments, including numbers, text, and logical values
* BETADIST -- Returns the beta cumulative distribution function
* BETAINV -- Returns the inverse of the cumulative distribution function for a specified beta distribution
* BINOMDIST -- Returns the individual term binomial distribution probability
* CHIDIST -- Returns the one-tailed probability of the chi-squared distribution
* CHIINV -- Returns the inverse of the one-tailed probability of the chi-squared distribution
* CHITEST -- Returns the test for independence
* CONFIDENCE -- Returns the confidence interval for a population mean
* CORREL -- Returns the correlation coefficient between two data sets
* COUNT -- Counts how many numbers are in the list of arguments -- _R_ _C_
* COUNTA -- Counts how many values are in the list of arguments -- _R_ _C_
* COUNTBLANK -- Counts the number of blank cells within a range
* COUNTIF -- Counts the number of nonblank cells within a range that meet the given criteria
* COVAR -- Returns covariance, the average of the products of paired deviations
* CRITBINOM -- Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value
* DEVSQ -- Returns the sum of squares of deviations
* EXPONDIST -- Returns the exponential distribution
* FDIST -- Returns the F probability distribution
* FINV -- Returns the inverse of the F probability distribution
* FISHER -- Returns the Fisher transformation
* FISHERINV -- Returns the inverse of the Fisher transformation
* FORECAST -- Returns a value along a linear trend
* FREQUENCY -- Returns a frequency distribution as a vertical array
* FTEST -- Returns the result of an F-test
* GAMMADIST -- Returns the gamma distribution
* GAMMAINV -- Returns the inverse of the gamma cumulative distribution
* GAMMALN -- Returns the natural logarithm of the gamma function, Γ(x)
* GEOMEAN -- Returns the geometric mean
* GROWTH -- Returns values along an exponential trend
* HARMEAN -- Returns the harmonic mean
* HYPGEOMDIST -- Returns the hypergeometric distribution
* INTERCEPT -- Returns the intercept of the linear regression line
* KURT -- Returns the kurtosis of a data set
* LARGE -- Returns the k-th largest value in a data set -- _R_ _C_
* LINEST -- Returns the parameters of a linear trend
* LOGEST -- Returns the parameters of an exponential trend
* LOGINV -- Returns the inverse of the lognormal distribution
* LOGNORMDIST -- Returns the cumulative lognormal distribution
* MAX -- Returns the maximum value in a list of arguments -- _R_ _C_
* MAXA -- Returns the maximum value in a list of arguments, including numbers, text, and logical values
* MEDIAN -- Returns the median of the given numbers
* MIN -- Returns the minimum value in a list of arguments -- _R_ _C_
* MINA -- Returns the smallest value in a list of arguments, including numbers, text, and logical values
* MODE -- Returns the most common value in a data set
* NEGBINOMDIST -- Returns the negative binomial distribution
* NORMDIST -- Returns the normal cumulative distribution
* NORMINV -- Returns the inverse of the normal cumulative distribution
* NORMSDIST -- Returns the standard normal cumulative distribution
* NORMSINV -- Returns the inverse of the standard normal cumulative distribution
* PEARSON -- Returns the Pearson product moment correlation coefficient
* PERCENTILE -- Returns the k-th percentile of values in a range
* PERCENTRANK -- Returns the percentage rank of a value in a data set
* PERMUT -- Returns the number of permutations for a given number of objects
* POISSON -- Returns the Poisson distribution
* PROB -- Returns the probability that values in a range are between two limits
* QUARTILE -- Returns the quartile of a data set
* RANK -- Returns the rank of a number in a list of numbers -- _R_ _C_
* RSQ -- Returns the square of the Pearson product moment correlation coefficient
* SKEW -- Returns the skewness of a distribution
* SLOPE -- Returns the slope of the linear regression line
* SMALL -- Returns the k-th smallest value in a data set
* STANDARDIZE -- Returns a normalized value
* STDEV -- Estimates standard deviation based on a sample
* STDEVA -- Estimates standard deviation based on a sample, including numbers, text, and logical values
* STDEVP -- Calculates standard deviation based on the entire population
* STDEVPA -- Calculates standard deviation based on the entire population, including numbers, text, and logical values
* STEYX -- Returns the standard error of the predicted y-value for each x in the regression
* TDIST -- Returns the Student's t-distribution
* TINV -- Returns the inverse of the Student's t-distribution
* TREND -- Returns values along a linear trend
* TRIMMEAN -- Returns the mean of the interior of a data set
* TTEST -- Returns the probability associated with a Student's t-test
* VAR -- Estimates variance based on a sample
* VARA -- Estimates variance based on a sample, including numbers, text, and logical values
* VARP -- Calculates variance based on the entire population
* VARPA -- Calculates variance based on the entire population, including numbers, text, and logical values
* WEIBULL -- Returns the Weibull distribution
* ZTEST -- Returns the one-tailed probability-value of a z-test

## Text functions

* ASC -- Changes full-width (double-byte) English letters or katakana within a character string to half-width (single-byte) characters
* BAHTTEXT -- Converts a number to text, using the ß (baht) currency format
* CHAR -- Returns the character specified by the code number
* CLEAN -- Removes all nonprintable characters from text
* CODE -- Returns a numeric code for the first character in a text string
* CONCATENATE -- Joins several text items into one text item -- _R_ _C_
* DOLLAR -- Converts a number to text, using the $ (dollar) currency format
* EXACT -- Checks to see if two text values are identical
* FIND, FINDB -- Finds one text value within another (case-sensitive) -- _R_ _C_ (not FINDB)
* FIXED -- Formats a number as text with a fixed number of decimals
* JIS -- Changes half-width (single-byte) English letters or katakana within a character string to full-width (double-byte) characters
* LEFT, LEFTB -- Returns the leftmost characters from a text value -- _R_ _C_ (not LEFTB)
* LEN, LENB -- Returns the number of characters in a text string -- _R_ (not LENB)
* LOWER -- Converts text to lowercase
* MID, MIDB -- Returns a specific number of characters from a text string starting at the position you specify - _R_
* PHONETIC -- Extracts the phonetic (furigana) characters from a text string
* PROPER -- Capitalizes the first letter in each word of a text value
* REPLACE, REPLACEB -- Replaces characters within text
* REPT -- Repeats text a given number of times
* RIGHT, RIGHTB -- Returns the rightmost characters from a text value -- _R_
* SEARCH, SEARCHB -- Finds one text value within another (not case-sensitive)
* SUBSTITUTE -- Substitutes new text for old text in a text string -- _R_
* T -- Converts its arguments to text
* TEXT -- Formats a number and converts it to text
* TRIM -- Removes spaces from text - _R_
* UPPER -- Converts text to uppercase
* VALUE -- Converts a text argument to a number

## External functions

* EUROCONVERT -- Converts a number to euros, converts a number from euros to a euro member currency, or converts a number from one euro member currency to another by using the euro as an intermediary (triangulation)
* SQL.REQUEST -- Connects with an external data source and runs a query from a worksheet, then returns the result as an array without the need for macro programming
