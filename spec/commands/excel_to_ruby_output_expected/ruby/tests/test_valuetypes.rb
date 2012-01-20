# Test for ValueTypes
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestValuetypes < Test::Unit::TestCase
  def worksheet; Valuetypes.new; end
end
end
