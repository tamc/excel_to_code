# Test for Referencing
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestReferencing < Test::Unit::TestCase
  def worksheet; Referencing.new; end
  def test_a1; assert_equal("Named reference",worksheet.a1); end
  def test_a2; assert_equal("Named reference",worksheet.a2); end
end
end
