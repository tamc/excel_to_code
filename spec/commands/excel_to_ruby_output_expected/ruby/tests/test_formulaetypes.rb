# Test for FormulaeTypes
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestFormulaetypes < Test::Unit::TestCase
  def worksheet; Formulaetypes.new; end
  def test_a1; assert_equal("Simple",worksheet.a1); end
  def test_b1; assert_equal(2,worksheet.b1); end
  def test_a2; assert_equal("Sharing",worksheet.a2); end
  def test_b2; assert_equal(267.7467614837482,worksheet.b2); end
  def test_a3; assert_equal("Shared",worksheet.a3); end
  def test_b3; assert_equal(267.7467614837482,worksheet.b3); end
  def test_a4; assert_equal("Shared",worksheet.a4); end
  def test_b4; assert_equal(267.7467614837482,worksheet.b4); end
  def test_a5; assert_equal("Array (single)",worksheet.a5); end
  def test_b5; assert_equal(2,worksheet.b5); end
  def test_a6; assert_equal("Arraying (multiple)",worksheet.a6); end
  def test_b6; assert_equal("Not Eight",worksheet.b6); end
  def test_a7; assert_equal("Arrayed (multiple)",worksheet.a7); end
  def test_b7; assert_equal("Not Eight",worksheet.b7); end
  def test_a8; assert_equal("Arrayed (multiple)",worksheet.a8); end
  def test_b8; assert_equal("Not Eight",worksheet.b8); end
end
end
