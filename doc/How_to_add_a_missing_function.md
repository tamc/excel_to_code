# How to add a missing function

Only some excel functions are currently implemented. These are listed in src/compile/ruby/map_formulae_to_ruby.rb in the MapFormulaeToRuby::FUNCTIONS hash.

## Adding a new function to the Ruby output

A shell script: util/add_function <function_name> will create the required files. 

Then fill out spec/excel/excel_functions/<function_name>_spec.rb with the appropriate specificaiton.

Then fill out src/excel/excel_functions/<function_name>.rb with the appropriate implementation.

Check that src/rewrite/expand_array_formulae_ast.rb will treat the formula correctly when it is given in {array} form. By default it assumes that none of the functions expect ranges (e.g., it is like {ABS(A1:A3)} and not like {SUM(A1:A3)} or {INDEX(A1:A3,B1:B3)})

## Adding a new function to the C output

Add the function definition and implementation to src/compile/c/excel_to_c_runtime.c

The tests for that function go in the main body of that file.

Run the tests with

    gcc excel_to_c_runtime.c; ./a.out

Add the function lookup to the FUNCTIONS hash in src/compile/c/map_formulae_to_c.rb

Note that if a function can vary in its number of arguments, then the map_formulae_to_c.rb code has approaches to call different C functions based on the number of arguments. 
