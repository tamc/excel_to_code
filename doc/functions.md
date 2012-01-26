= Functions

Only some excel functions are currently implemented. These are listed in src/compile/ruby/map_formulae_to_ruby.rb in the MapFormulaeToRuby::FUNCTIONS hash.

== Adding a new function

A shell script: util/add_function <function_name> will create the required files. 

Then fill out spec/excel/excel_functions/<function_name>_spec.rb with the appropriate specificaiton.

Then fill out src/excel/excel_functions/<function_name>.rb with the appropriate implementation.