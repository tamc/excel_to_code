# coding: utf-8
# Test for (innapropriate) sheet name!
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestS_innapropriate_sheet_name_ < Test::Unit::TestCase
  def worksheet; S_innapropriate_sheet_name_.new; end
  def test_c4; assert_in_epsilon(1,worksheet.c4); end
end
end
