# Test for FormulaeTypes
require 'test/unit'
require_relative 'formulaetypes'

class TestFormulaetypes < Test::Unit::TestCase
  def test_a1; assert_equal(worksheet.a1,"Simple"); end
  def test_b1; assert_equal(worksheet.b1,2); end
  def test_a2; assert_equal(worksheet.a2,"Sharing"); end
  def test_b2; assert_equal(worksheet.b2,267.7467614837482); end
  def test_a3; assert_equal(worksheet.a3,"Shared"); end
  def test_b3; assert_equal(worksheet.b3,267.7467614837482); end
  def test_a4; assert_equal(worksheet.a4,"Shared"); end
  def test_b4; assert_equal(worksheet.b4,267.7467614837482); end
  def test_a5; assert_equal(worksheet.a5,"Array (single)"); end
  def test_b5; assert_equal(worksheet.b5,2); end
  def test_a6; assert_equal(worksheet.a6,"Arraying (multiple)"); end
  def test_b6; assert_equal(worksheet.b6,"Not Eight"); end
  def test_a7; assert_equal(worksheet.a7,"Arrayed (multiple)"); end
  def test_b7; assert_equal(worksheet.b7,"Not Eight"); end
  def test_a8; assert_equal(worksheet.a8,"Arrayed (multiple)"); end
  def test_b8; assert_equal(worksheet.b8,"Not Eight"); end
end
