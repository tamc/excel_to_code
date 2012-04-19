# All tests for /Users/tamc/Documents/github/excel_to_code/spec/test_data/ExampleSpreadsheet.xlsx
require 'test/unit'
require_relative 'examplespreadsheet'

class TestExamplespreadsheet < Test::Unit::TestCase
  def worksheet; @worksheet ||= Examplespreadsheet.new; end
  # Start of ValueTypes
  # End of ValueTypes

  # Start of FormulaeTypes
  # End of FormulaeTypes

  # Start of Ranges
  # End of Ranges

  # Start of Referencing
  # End of Referencing

  # Start of Tables
  # End of Tables

  # Start of (innapropriate) sheet name!
  # End of (innapropriate) sheet name!

end
