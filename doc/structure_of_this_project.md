# Excel_to_code project structure

* bin - empty at the moment
* doc - notes and documentation for improving the code
* examples - empty at the moment
* spec - contains the tests that make sure the code is working properly
    * spec/test_data - contains bits of xml and other that are used in the testing
    * spec_helper - contains global methods and setup, available to all tests
* benchmark - contains some tests of how fast some bits of code are, useful when trying to optimise
* src - contains the actual code used to run the project
    * src/commands - contains the higher level code that executes commands, such as converting an entire workbook into ruby
    * src/compile - contians code that takes the rewritten text files and converts them into code (e.g., a ruby version of a spreadsheet) 
    * src/compile/ruby - contains code that takes the rewritten text files and converts them into ruby
    * src/compile/c - contains code that takes the rewritten text files and converts them into c
    * src/excel - contains code that manipulates and processes excel formulae, such as the formula parsers and classes for manipulating references
    * src/excel/functions - contains ruby implementations of each of the supported excel functions
    * src/extract - contains code that takes the Excel xml and converts it into a series of text files
    * src/rewrite - contains code that takes the series of text files and converts them into different text files (for instance, to remove INDIRECT functions, or to turn the excel form of formulas into ast)
    * src/simplify - contains code that takes the rewritten text files and tries to simplify the calculations by, for instance, replacing functions with their values if they can be calculated at runtime
    * src/util - contians classes that aren't used at runtime but that might be helpful in development
