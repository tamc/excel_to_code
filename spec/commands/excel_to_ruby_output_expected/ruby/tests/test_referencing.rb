# Test for Referencing
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestReferencing < Test::Unit::TestCase
  def worksheet; Referencing.new; end
  def test_a1; assert_equal(worksheet.a1,"Named reference"); end
  def test_a2; assert_equal(worksheet.a2,"Named reference"); end
end
end
