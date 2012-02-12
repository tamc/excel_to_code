# coding: utf-8
# Test for Tables
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestTables < Test::Unit::TestCase
  def worksheet; Tables.new; end
  def test_a1; assert_in_epsilon(12,worksheet.a1); end
end
end
