# coding: utf-8
# Test for Referencing
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestReferencing < Test::Unit::TestCase
  def worksheet; Referencing.new; end
  def test_a4; assert_in_epsilon(10,worksheet.a4); end
end
end
