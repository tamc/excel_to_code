# Parsing excel formulae

The parser that this project uses to turn excel formulae into abstract syntax trees is given in src/excel/formula_peg.txt. This is turned into the ruby code in formula_peg.rb using the command:

	cd src/excel
	text-peg2ruby-peg formula_peg.txt formula_peg.rb

The form of the grammar is given in the docs relating to the rubypeg gem, available from https://github.com/tamc/rubypeg

## Testing the parser

Tests can be added to spec/test_data/formulae_to_ast.txt

In that file, each line contains an excel formula, followed by the ast into which it should be converted.

The tests are run by:

    rspec spec/excel/formula_peg_spec.rb

## Probable bugs in parsing

* Whitespace handling in table references
* Whitespace handling between arguments in a list