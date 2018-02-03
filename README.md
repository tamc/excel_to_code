# Excel to Code

[![Tests Passing](https://travis-ci.org/tamc/excel_to_code.svg?branch=master)](https://travis-ci.org/tamc/excel_to_code)

excel_to_c - roughly translates some Excel files into C.

excel_to_ruby - roughly translates some Excel files into Ruby.

The motivation is to allow spreadsheets to be:

1. Embedded in other programs (such as web servers, or optimisers)
2. Without depending on any Microsoft library (the generated files are self contained)

# Install

Requires Ruby. Install by:

    gem install excel_to_code

# Run

To just have a go:

	./bin/excel_to_c <excel_file_name>
	

For a more complex spreadsheet:
	
	./bin/excel_to_c --compile --run-tests --settable <name of input worksheet> --prune-except <name of output worksheet> <excel file name> 
	
See the full list of options:

	./bin/excel_to_c --help


# Report bugs and fixes

Source code: [https://github.com/tamc/excel_to_code]

Report bugs: [https://github.com/tamc/excel_to_code/issues]

# Hacking excel_to_code

There are some how to guides in the doc folder. 

## Test

1. Make sure you have ruby 1.9.2 or later installed
2. gem install bundler # May need to use sudo
3. bundle
4. rspec spec/*

To test the C runtime:
1. cd src/compile/c
2. cc excel_to_c_runtime
3. ./a.out

# Limitations

1. Not tested at all on Windows
2. It must be possible to replace INDIRECT and OFFSET formula with standard references at compile time (e.g., INDIRECT("A"&"1") is fine, INDIRECT(userInput&"3") is not.
3. Doesn't implement all functions (see doc/Which_functions_are_implemented.md)
4. Doesn't implement references that involve range unions and lists (but does implement standard ranges)
5. Sometimes gives cells as being empty, when excel would give the cell as having a numeric value of zero
6. The generated C version does not multithread and will give bad results if you try
7. The generated code uses floating point, rather than fully precise arithmetic, so results can differ slightly
8. The generated code uses the sprintf approach to rounding (even-odd) rather than excel's 0.5 rounds away from zero.
90. Ranges like this: Sheet1!A10:Sheet1!B20 and 3D ranges don't work
