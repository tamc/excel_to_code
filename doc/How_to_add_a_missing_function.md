# How to add a missing function

Only some excel functions are currently implemented. These are listed in src/compile/ruby/map_formulae_to_ruby.rb in the MapFormulaeToRuby::FUNCTIONS hash.

== Adding a new function

A shell script: util/add_function <function_name> will create the required files. 

Then fill out spec/excel/excel_functions/<function_name>_spec.rb with the appropriate specificaiton.

Then fill out src/excel/excel_functions/<function_name>.rb with the appropriate implementation.

Check that src/rewrite/expand_array_formulae_ast.rb will treat the formula correctly when it is given in {array} form. By default it assumes that none of the functions except ranges (e.g., it is like {ABS(A1:A3)} and not like {SUM(A1:A3)} or {INDEX(A1:A3,B1:B3)})