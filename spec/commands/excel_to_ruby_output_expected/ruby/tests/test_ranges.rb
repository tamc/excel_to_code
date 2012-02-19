# coding: utf-8
# Test for Ranges
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestRanges < Test::Unit::TestCase
  def spreadsheet; $spreadsheet ||= Spreadsheet.new; end
  def worksheet; @worksheet ||= spreadsheet.ranges; end
  def test_b1; assert_equal("This sheet",worksheet.b1); end
  def test_c1; assert_equal("Other sheet",worksheet.c1); end
  def test_a2; assert_equal("Standard",worksheet.a2); end
  def test_b2; assert_in_epsilon(6,worksheet.b2); end
  def test_c2; assert_in_epsilon(4.141500000000001,worksheet.c2); end
  def test_a3; assert_equal("Column",worksheet.a3); end
  def test_b3; assert_in_epsilon(6,worksheet.b3); end
  def test_c3; assert_equal(:name,worksheet.c3); end
  def test_a4; assert_equal("Row",worksheet.a4); end
  def test_b4; assert_in_epsilon(6,worksheet.b4); end
  def test_c4; assert_in_epsilon(3.1415,worksheet.c4); end
  def test_f4; assert_in_epsilon(1,worksheet.f4); end
  def test_e5; assert_in_epsilon(1,worksheet.e5); end
  def test_f5; assert_in_epsilon(2,worksheet.f5); end
  def test_g5; assert_in_epsilon(3,worksheet.g5); end
  def test_f6; assert_in_epsilon(3,worksheet.f6); end
end
end
