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
** src/extract - contains code that takes the Excel xml and converts it into a series of text files
** src/rewrite - contains code that takes the series of text files and converts them into different text files (for instance, to remove INDIRECT functions, or to turn the excel form of formulas into ast)
** src/excel - contains code that manipulates and processes excel formulae, such as the formula parsers and classes for manipulating references
** src/output - contians code that takes the rewritten text files and converts them into code (e.g., a ruby version of a spreadsheet)


# TODO

* Sort out parsing of unions and literal arrays
* Sort out extraction of sheet dimensions
* Sort out rewriting column and row references to bounded range references
* Sort out knowing what functions want as arguments in order to do the appropriate casting
