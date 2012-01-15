# Test for ValueTypes
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestValuetypes < Test::Unit::TestCase
  def worksheet; Valuetypes.new; end
  def test_a1; assert_equal(true,worksheet.a1); end
  def test_a2; assert_equal("Hello",worksheet.a2); end
  def test_a3; assert_equal(1,worksheet.a3); end
  def test_a4; assert_equal(3.1415,worksheet.a4); end
  def test_a5; assert_equal(:name,worksheet.a5); end
  def test_a6; assert_equal("Hello",worksheet.a6); end
end
end
