# Test for Referencing
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestReferencing < Test::Unit::TestCase
  def worksheet; Referencing.new; end
  def test_a1; assert_equal("Named reference",worksheet.a1); end
  def test_a2; assert_equal("Named reference",worksheet.a2); end
  def test_a4; assert_equal(10,worksheet.a4); end
  def test_b4; assert_equal(11,worksheet.b4); end
  def test_c4; assert_equal(12,worksheet.c4); end
  def test_a5; assert_equal(3,worksheet.a5); end
  def test_b8; assert_equal("Named reference",worksheet.b8); end
  def test_b9; assert_equal(3,worksheet.b9); end
end
end
