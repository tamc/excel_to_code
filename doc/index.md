# Excel2Code project structure

In brief:

* bin - empty at the moment
* doc - notes and documentation
* examples - empty at the moment
* spec - contains the tests that make sure the code is working properly
** spec/test_data - contains bits of xml that are used in the testing
** spec_helper - contains global methods and setup, available to all tests
* src - contains the actual code used to run the project
** src/commands - contains the higher level code that executes commands, such as converting an entire workbook into ruby
** src/compile - contians code that takes the rewritten text files and converts them into code (e.g., a ruby version of a spreadsheet) 
** src/excel - contains code that manipulates and processes excel formulae, such as the formula parsers and classes for manipulating references
** src/excel/functions - contains ruby implementations of each of the supported excel functions
** src/extract - contains code that takes the Excel xml and converts it into a series of text files
** src/rewrite - contains code that takes the series of text files and converts them into different text files (for instance, to remove INDIRECT functions, or to turn the excel form of formulas into ast)
** src/simplify - contains code that takes the rewritten text files and tries to simplify the calculations by, for instance, replacing functions with their values if they can be calculated at runtime
** src/util 

# TODO

* Sort out parsing of unions
* Add missing functions
* Add output in other programming languages

