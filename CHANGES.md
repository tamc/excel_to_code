# Head

# 0.3.19 - 2018 January 8

## Bug fixes

- Remove unused Object#try (thank you slisser)

# 0.3.18 - 2018 January 5 

## Bug fixes

- Implement the '#,000' format in Ruby and C versions of the TEXT function

# 0.3.18 beta 2 - 2018 December 30

## New / Changed Excel functions

- Implement MROUND in C and Ruby
- Implement SUBSTITUTE in C

## Bug fixes

- Make conditional criteria (e.g., ">10") in countif/sumif/countifs/sumifs robust to spaces (e.g., "> 10")
- Treat comparators as values when extracting common elements from formulae
- Fix replacement of range arguments passed to the OFFSET function
- Make OFFSET cope with references that have been inlined
- Add some missing error checks to C versions of functions
- Increase maximum number of simplifying passes from 20 to 50
- Make it harder for simplifications to accidentally trample on named reference setters

# 0.3.18 beta 1 - 2018 July 16

## New / Changed Excel functions

- Implement CEILING in ruby and C
- Add FORECAST.LINEAR as an alias of the FORECAST function
- Implement REPLACE in ruby (BUT NOT IN C)
- Partially implement RATE in ruby and C
- Implement COUNTIFS in ruby and C
- Implement PRODUCT in ruby and C
- Implement FLOOR in ruby and C
- Implement SQRT in ruby and C
- Add the OR() function in Ruby and C
- INDEX can now work with ranges passed as its second and/or third arguments
- Add the NA() function in Ruby and C
- Fix COLUMN() and ROW() to work on area, named and table references
- Replace CELL("address", REF) at compile time
- Add the HYPERLINK() function in Ruby
- Added the NOT(true) function in Ruby, and C

## Other new features

- Allow wildcards in ruby version of SUMIFS when string matching string
- Add allow_unknown_functions option
- Add treat_external_references_as local helper
- Add a debug_dump() function for de-bugging
- Add --persevere switch which will mean will not abort on parse or external reference error so you can see the scale of the problem in a workbook.
- Better error reporting on external references in named references
- Add a util/check.rb command to help trace translation errors

## Bug fixes

- Allow function names to have periods, be lowercase and have _
- Fix C version of MATCH to correctly skip strings when matching numbers and vice versa
- Fix require_relative in gempsec that was causing 'Bundler cannot continue' errors
- Make replacement of arrays with single cells depth first
- Fix bug in prefixes inside arithmetic applied to arrays
- Improve conversion of prefixes applied to arrays
- Improvements to the way translator converts range references that should be treated as cell references
- Indicate that rspec tests use both old and new syntax
- Fix a bug in range checking in the C version of the LARGE function
- Fix bug in keeping of table references when setting to keep all named references
- Fix so that works when there are apostrophes in the path [#12](https://github.com/tamc/excel_to_code/issues/12)
- Improve bracket elimination when simplifying arithmetic
- Fix IF functions where comparison is an array

## Refactors

- Add a string_argument helper function in Ruby
- Experiment with ways of making it easier to wrap the C output in a command line interface
- Update dependencies
- Refactor common command line options
- Reformat command line help

# 0.3.17 - 2015 May 14

- Fix a bug when simplifying SUM("Not a number") and the like

# 0.3.16 - 2015 May 12

- Be more permissive of newlines in strings and formulae
- Raise an Exception if fail to parse a formulae
- Generated Ruby tests now ignore differences in newlines in strings
- Implemented the ADDRESS function in Ruby
- PMT now accepts 3, 4 or 5 arguments. Arguments 4 or 5 must be zero at the moment.

# 0.3.15 - 2015 March 10

- Fixed a bug in VLOOKUP when the lookup value is an empty cell

# 0.3.14 - 2015 March 10

- Fixed bug in Ruby version of power
- Now detects and reports more kinds of external references
- Improved the way that the conversion treats newlines and other special characters in strings by dealing with excel unicode escaping of the form _x000D_
- Added missing NUM conversion to Ruby interface to C version
- Fixed PMT to return NUM error when number of periods is zero

# 0.3.13 - 2015 March 2

- Roots of negative numbers now return an error, like excel
- Fixed a bug when specifying :all cells in a sheet to be settable, and empty cells on that sheet are referenced

# 0.3.12 - 2015 February 27

- Fixed a bug which meant that COLUMN and ROW and OFFSET functions would sometimes not be correctly replaced
- Tinkered with tests

# 0.3.11 - 2015 February 6

- Implemented the CHAR function in C and Ruby
- Default logging now uses colours to indicate severity

# 0.3.10 - 2015 January 16

- Fix unitialzed constant error

# 0.3.9 - 2015 January 15

- Fix SUBTOTAL to not count other SUBTOTALs
- Tweak the logging of the initial unzip of the excel into xml to be quieter
- Altered the output of the default logger to be less noisy

# 0.3.8 - 2015 January 15

- Tiny wording tweak to the external reference reporting

# 0.3.7 - 2015 January 15

- Reports more detail when it finds an external reference
- Implemented the ISERROR() function

# 0.3.6 - 2014 December 19th

- Fixed a stupid bug causing it always writes the tests in C

# 0.3.5 - 2014 December 18th

- Fixed a bug in the C version of string_join when joining some strings with more than 100 characters
- Generated C file now uses EXCEL_STRING, EXCEL_NUMBER and EXCEL_RANGE macro names instead of more confusing new_excel_string, new_excel_number and new_excel_range
- Now replaces values with constants in named references
- Fixed some memory leaks in the C runtime
- Fixed bug in RIGHT() and LEFT() which would give empty strings when number of characters requested was greater than the size of the string passed
- Can now generate tests in pure C as well as in Ruby using a new --write-tests-in-c option
- If choose to generate tests in pure C, will still produce Ruby version

# 0.3.4 - 2014 December 1st

- Fixed bug when output directory has been set and option set to compile and test C code
- Implemented the NPV function in Ruby and C
- Copes when named references contain table references

# 0.3.3 - 2014 October 7th

- Ruby version of FIND function now works when arguments are not strings

# 0.3.2 - 2014 July 28th

- Replaced Nokogiri with Ox

# 0.3.1 - 2014 July 25th

- Fixed so that compiling and testing works even if --no-rakefile is specified

# 0.3.0 - 2014 July 24th

- Shim extension not needed on generated Ruby to C interface code (BREAKING CHANGE), which means that excel_to_c generates a more-or-less drop in replacement for excel_to_Ruby
- Added --rakefile and ---makefile and --no-rakefile and --no-makefile switches to excel_to_c
- By default, now generates a Rakefile (BREAKING CHANGE)

# 0.2.30 - 2014 July 24th

- Implemented 0.000 in C version of TEXT function (already works in Ruby version)

# 0.2.29 - 2014 July 2nd

- Now replaces table references that have been interpreted as named references
- Implemented 0.0 and 0000 and #,## as formats in the TEXT function
- RIGHT and LEFT can now be applied to an array
- Fixed INDEX when passed zero as a row or column number when only one row or column wide
- Implemented the ISERR() function in Ruby and in C
- Implmeneted the LOG10() function

# 0.2.28 - 2014 June 16th

- Fixed a bug that would lead to some numbers being truncated to integers in the C version of the output

# 0.2.27 - 2014 June 16th

- named_references_to_keep and named_references_that_can_be_set_at_runtime can now refer to table names

# 0.2.26 - 2014 June 10th

- When pruning cells, now checks for circular references
- Added a better error message if expanding array formulae fails
- Now doesn't mistakenly inline sheet specific named references
- Made the array formula expansion more robust to mismatched array sizes

# 0.2.25 - 2014 May 23rd

- Implemented the LN function in Ruby and C
- Tweaks to the array formulae code to try and speed it up
- Tightened dependency version numbers in gemspec to eliminate warnings when building gem

# 0.2.24 - 2014 May 16th

- Fixed the simplification of SUMIFS when it has errors as arguments

# 0.2.23 - 2014 May 7th

- Fixed the simplification of SUMIFS when a criteria is blank
- Fixed the conversion of INDEX formulae that return arrays when they are inside array formulae

# 0.2.22 - 2014 April 27th

- Implemented a reset method in the Ruby code so can perform repeated calculations on the same instance (already implemented in the generated C code)
- Added named references to generated Ruby code (already implemented in generated C code)
- Added settting to allow the inline_formulae_that_are_only_used_once method to be skipped

# 0.2.21 - 2014 April 27th

- Fix Ruby version of SUMIFS to match number criteria against string check values (already fixed in C version)
- bin/excel_to_c and bin/excel_to_Ruby options changed and added -n option to keep named references
- default behaviour when named_references_to_keep = :all but no named references is then to keep all cells

# 0.2.20 - 2014 March 31st

- Attempts to simplify some SUMIFS statements, even if they can't be calculated entirely at compile time
- Removed redundant ReplaceFormulaeWithCalculatedValues class and test - The underlying MapFormulaeToValues class is now used directly
- Fixed comparisons to be able to compare values of different types (e.g., true is always greater than String in Excel)
- Added an ExcelToTest class which generates a test file for a workbook, but does not generate the code to implement the workbook.

# 0.2.19 - 2014 March 18th

- Fixed a bug in ReplaceArithmeticOnRangesAst where there was too much mapping going on
- Improved the number of cases where SUM is partially eliminated
- Now replaces IF statements at compile time if the condition can be evaluated
- Stop replacing 0^X and 0/X with zero. If X is zero, then an error is the correct answer.
- Inlined blanks are now returned as blanks not at zeros.
- When extracting constants, now uses the integer numbers from one to ten that are defined as constants in the C runtime
- Improved the annoying emergency bodge on array formulae extraction to interpret INDIRECTs within array formulae in more cases
- Can now parse some formulae with a trailing percentage suffix like ROUND(24.3,0)%
- If INDEX is passed an errors in both arguments, returns the error in its first argument
- Now simplifies SUM() with single arguments to the single argument
- Now deals with arrays that should be converted to single cells inside of IF statements
- Altered how SUMIFS deals with passed errors in Ruby verison (need to check C version)
- Can now partially replace SUMIFS in situations where it is being used as a form of lookup
- Refactored AVERAGEIFS and SUMIFS in Ruby to be more like C
- Better error message when named reference can't be parsed
- Changed the default match_type in the MATCH function to be 1
- Fixed a bug when a table has column names containing a # symbol
- Can now isolate more than one worksheet at the same time

# 0.2.18 - 2014 March 2nd

- Had another go at dealing with the way that blank cells become zeros when referenced in some ways but not others
- Named references that are actually #REF! errors are correctly parsed and then turned into errors
- Implemented the VALUE function

# 0.2.17 - 2014 February 11th

- Changed how SUMIF, SUMIFS, and AVERAGEIFS treat blanks in criteria (now consider them to be zeros)
- Add missing cells if they are referenced
- Adjust the way that ISBLANK() is replaced with a value at compile time

# 0.2.16 - 2014 February 10th

- Fixed bug in response when LEFT and RIGHT are passed negative values
- Implemented LEN function in C runtime
- Implemented RIGHT funtion in C runtime
- Inline blanks are treated as empty strings in MID and RIGHT functions
- Fixed a bug in bounds checking when replacing arithmetic on ranges
- Replace comparisons on arrays with arrays of comparisons
- Ruby version of RIGHT now returns a string if index out of bounds
- Removed the special case optimisation of the COUNT function. It wasn't right in some cases
- Now copes with passing a range where a cell expected in the CHOOSE function, but only if CHOOSE is the first function in a cell
- Now copes with a wider variety of situations where a range that should be converted into a cell is passed to a match or indirect function
- When simplifying arithmetic, ensures results are always numbers
- Added an ENSURE_IS_NUMBER function to the runtime which turns anything that Excel thinks looks a bit like a number into a real number
- Added an inlined_blank type, which is sometimes like a zero, and sometimes like an empty string
- String join in the Ruby compilation treates zeros that are blank differently from zeros
- Added an isolate attribute to excel_to_x to help debugging large worksheets: if set, only one sheet is translated
- Now prints more informative error messages if fails to compile a 'common' element
- Changed order of array replacement
- Ruby version of LEN now distinguishes zeros that are blank from zeros that are not
- References to blank cells now return 0 not blank
- Fix for SUMIF different sized ranges when fixed references are used
- Parser now picks up external named references
- Added a workaround for brackets not being removed in some cases
- The Ruby version of IF() now treats 0 as false and all other numbers as true
- SUMIF and SUMIFS can now cope with being passed ranges for the criteria
- Fixed formula parser to cope with comparisons with string joins
- Fixed bug in C compilation when specifying :all cells in a sheet are settable
- Switched off warnings in default C build script

# 0.2.15 - 2014 February 6th

- SUMIF now copes with criteria ranges that are different in size from sum ranges
- Optional attribute on excel_to_x: extract_repeated_parts_of_formulae. Setting it to false may make it easier to localise errors
- Generated C code now has C if statements, so that there is no need to evaluate both the true and the false case
- OFFSET replacement now deals with errors in arguments
- Reworked the way that ranges are dealt with when cells are expected (may not be right)

# 0.2.14 - 2014 February 4th

- Optimised the array formula expansion a tiny bit
- Aborts earlier if unable to replace a ROW or COLUMN function
- When inlining, don't inline into a COLUMN or ROW function

# 0.2.13 - 2014 February 3rd

- Now fails earlier if external references are found
- Extra logging messages when expanding array formulae

# 0.2.12 - 2014 February 3rd

- Reworked the SimplifyArithmetic routine to make fewer passes.

# 0.2.11 - 2014 February 2nd

- Implemented the FORECAST function
- Implemented the AVERAGEIFS function
- Implemented ExcelToCode.version and display this when running the program or listing errors
- Updated how to add a missing function docs

# 0.2.10 - 2014 February 1st

- Now aborts earlier in the process if the spreadsheet contains functions that haven't yet been implemented in excel_to_code

# 0.2.9 - 2014 January 31st

- implemented the TRANSPOSE function (removed at compile time, only works in array formulae)

# 0.2.8 - 2014 January 31st

- implemented the ISBLANK function in Ruby and C

# 0.2.7 - 2014 January 30th

- implemented the EXP function in Ruby and C

# 0.2.6 - 2014 January 30th

- implemented the ROW() function (so long as it can be replaced at runtime)
- implemented the COLUMN() function without arguments

# 0.2.5 - 2014 January 13th

- named_references_to_keep and named_references_that_can_be_set_at_runtime can now be passed a block which will be called with each named reference. Returning true will make the named reference kept / settable

# 0.2.4 - 2014 January 13th

- Now checks whether sheet names and named references provided by the user actually exist, and if not, aborts

# 0.2.3 - 2014 January 4th

- Implemented the LOWER function in Ruby
- Removed the run_in_memory option, it no longer does anything.
- Remmoved the intermediate_directory option, the code no longer needs to write intermediate files

# 0.2.2 - 2013 December 29th

- Moved the excel function tests out of the C runtime and into a separate file
- Moved the reset() function into the C runtime out of the generated code
- Changed the generated tests to use minitest, now it is part of the Ruby library
- Removed the ZenTest dependency
- Changed the C runtime to avoid the big static arrays, and instead malloc the amount of memory actually needed

# 0.2.1 - 2013 December 22nd

- Fixed a bug when named references are specified as well as cells to keep and named reference points to blanks

# 0.2.0 - 2013 December 20th

- A lot of changes: no longer generates intermediate files, hopefully faster, but perhaps not
- Uses Ox instead of Nokogiri to parse worksheets, and does the parsing in one pass.
- Fixed a bug in the Ruby version of vlookup and hlookup that would downcase strings passed as arguments in certain situations

# 0.1.23 - 2013 December 4th

- Implemented the RIGHT function in Ruby

# 0.1.22 - 2013 December 4th

- Fixed a bug in = and != when comparing a string with a non string
- Implemented the LEN function in Ruby
- Implemented the SUBSTITUTE function in Ruby

# 0.1.21 - 2013 November 19th

- String joins can now be done on excel ranges
- Fixed bug in parsing of table references when they look like this: table.name[[column_name_inside_two_square_brackets]]
- Fixed OFFSET to work with sheet as well as cell references
- Implemented ISNUMBER function
- Fixed bug in OFFSET replacement
- Can now parse structured table references like this: G.40.levels.efficiency[[Description]:[Description]]
- Fixed StringIO#lines is deprecated warning.

# 0.1.20 - 2013 August 28th

- Fix compiler warnings relating to comparing a string literal

# 0.1.19 - 2013 August 24th

- Implemented the RANK() function

# 0.1.18 - 2013 August 23rd

- Implemented the CONCATENATE() function

# 0.1.17 - 2013 August 22nd

- Changed the way that shared targets are converted into a hash to avoid a stack overflow when there are very large numbers of shared formulae in a worksheet.

# 0.1.16 - 2013 August 21st

- Now returns the correct answer when MMULT() is used outside of an array formula

# 0.1.15 - 2013 August 21st

- Now translates formulae that peform arithmetic on ranges, such as SUMPRODUCT(A1:A10, 1/B1:B10)

# 0.1.14 - 2013 August 20th

- Implementing the MMULT function in Ruby and C (note only works when entered as an array formula)

# 0.1.13 - 2013 August 19th

- Implemented LOG function in Ruby and C
- Now converts the POWER(10,2) function as well as 10^2

# 0.1.12

- Fixed bug in working out how many passes to carry out when trying to eliminate INDIRECT() and other functions

# 0.1.11

- C version Now returns NUM error when trying to do a non integer power of a negative number
- Previously where shared formulae overlapped, the last would win. Now the correct one wins.
- Tests are now output in dependency order, so the first test to fail is likely to be closest to the underlying problem.
- VLOOKUP and HLOOKUP now accept 1 and 0 in place of TRUE and FALSE on their last argument
- Inline strings now have newlines stripped to be consistent with shared strings (it would be better if they were stripped in neither)
- INDEX now returns zero, not blank, when pointing to an empty cell
- If a range points to a single cell, then that is used rather than a range
- Fixed a bug where a nested sum would return 0 rather than an error in C code
- Implementing the HLOOKUP function
- When running in memory mode, now allows old intermediate files to be garbage collected, reducing memory use

# 0.1.10

- Now replaces INDIRECT() functions that have a second argument (but note, excel_to_code only supports A1 refs, not R1C1).
- Now repeats its passes to replace formulae with values until there is no more to be done
- Partially implemented the TEXT() function (for '0%' format only at this stage)
- Can now replace INDIRECT() functions where they argument is an error
- Inlining now copes with references to missing worksheets
- Implement the PV() function in Ruby and C
- Implemented TRIM(), MID() and partially implemented CELL() in Ruby
- The negative prefix can now operate on ranges
- Now attempts to simplify COLUMN() functions out of the spreadsheet at compile time (the function is not supported at runtime)
- Fix bug in the treatment of array results when simplifying INDEX functions
- Fixed a bug in replacing some table references with their values
- Now attempts to simplify OFFSET() functions out of the spreadsheet at compile time (the function is NOT supported at runtime)
- COUNT() functions are now replaced with their values if possible at compile time
- New option for cells_that_can_be_set_at_runtime = :named_references_only that only creates settable values for references that have been set in named_references_that_can_be_set_at_runtime.
- Fixed a bug in array formulae expansion: now copes with functions that have a variable number of arguments

# 0.1.9

- Removed chartsheets from the list of worksheets

# 0.1.8

- Remove newlines from excel formulae. This may turn out to be a bad idea (newlines in strings in formulae?) but needed for now.
- Implemented the LARGE(array,k) excel function in C and in Ruby

# 0.1.7

- Implemented the INT(number) excel function in C and in Ruby.

# 0.1.6

- named_references_to_keep can now be set to as :all to create getters for all named references
- named_references_that_can_be_set_at_runtime can now be set to :where_possible to create setters for named references that point to settable cells

# 0.1.5

- map_formulae_to_values now caches its results, which dramatically speeds up some edge-case long formulae
- Fixed an error in the way the C version of excel's IF function was implemented. It no longer returns an error if the unused argument is an error.
- Fixed an error in the way that the Ruby shim for the C version handled functions that return ranges (appears when trying to access named reference ranges)
- Now transfer named references into cells to keep early in the process so that referenced but empty cells are kept in final results

# 0.1.4

- Changed from requiring Ruby 1.9 to requiring 1.9 or newer.

# 0.1.3

- Closed #9 a bug which left Getsetranges where it shouldn't be.
- Closed #10 a bug that gave named references a prefix of '_'

# 0.1.2

- C version: Fixed naming of common methods
- Updated the way that command tests are run
- Merged energynumbers memory freeing code

# 0.1.1

- C version: now optionally writes out accessors for named references
- C version: shim can now get and set arrays of values, where the underlying C code supports it (which is only the case for some named references)
- C version: fixed bug in mapping of sheet names that could occur if two sheets had similar names

# 0.1.0

- BACKWARDS INCOMPATIBLE CHANGE: The ExcelValue struct in the Ruby FFI interface that is generated when compiling excel to C now gives type 'pointer' to its 'string' attribute. This is so that string values can be written through the FFI interface as well as read.

- C version now generates a Shim class that automatically translates between Ruby values and excel values. This makes it closer to being drop in compatible with the Ruby version.

# 0.0.14

- Fix parsing of non-western characters in formulae and named references

# 0.0.13

- Fix homepage in gem

# 0.0.12

- By default, the generated tests for the generated code are more relaxed about how closely numbers match.
- Command line option to generate tests that require an exact match

# 0.0.11

- Better handling of shared formulae in Excel: copes with cells that are exceptions to the general sharing rule

# 0.0.10

- Increased the default heap size for the memory to be freed later heap

# 0.0.9

- Memory used by generated C code can now be freed, reducing memory leak

# 0.0.8

- Simplified the code by making it automatically generate filenames for intermediate files
- Added a script to clean up the examples directory
- Fix bug that removed too many cells if the user hadn't specified the cells that they wished to keep

# 0.0.7

- The dynamic library name is no longer hardwired in the generated C makefile
- Better defaults for which cells are considered settable

# 0.0.6

- Ensure that if you have explicitly made a cell settable, it always appears in the resulting, even if it is blank or unneeded by the output functions

# 0.0.5

- Fixed intermittent bug in average function of excel_to_c_runtime.c

# 0.0.4

- Specifing an output name in snake_case will now cause camel case Ruby module names to be created (e.g., --output-name simple_model causes a class called SimpleModel to be created )
- Fixed bug where rubypeg dependency was not specified.

# 0.0.3

- Refactoring
- Fixed bug when specifying individual cell dependencies

# 0.0.2

Added an option to just keep intermediate files in memory rather than writing them to disk

# 0.0.1

First release as a gem
