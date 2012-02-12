# coding: utf-8
# Test for Ranges
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestRanges < Test::Unit::TestCase
  def worksheet; Ranges.new; end
end
end
