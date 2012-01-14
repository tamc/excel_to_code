# Test for ValueTypes
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestValuetypes < Test::Unit::TestCase
  def worksheet; Valuetypes.new; end
  def test_a1; assert_equal(worksheet.a1,true); end
  def test_a2; assert_equal(worksheet.a2,"Hello"); end
  def test_a3; assert_equal(worksheet.a3,1); end
  def test_a4; assert_equal(worksheet.a4,3.1415); end
  def test_a5; assert_equal(worksheet.a5,:name); end
  def test_a6; assert_equal(worksheet.a6,"Hello"); end
end
end
