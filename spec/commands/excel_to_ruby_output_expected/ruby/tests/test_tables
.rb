# Test for Tables
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestTables
 < Test::Unit::TestCase
  def worksheet; Tables
.new; end
  def test_a1; assert_equal(12,worksheet.a1); end
end
end
