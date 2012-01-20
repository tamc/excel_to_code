# Test for FormulaeTypes
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestFormulaetypes < Test::Unit::TestCase
  def worksheet; Formulaetypes.new; end
end
end
