# coding: utf-8
# Test for (innapropriate) sheet name!
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestS_innapropriate_sheet_name_ < Test::Unit::TestCase
  def spreadsheet; $spreadsheet ||= Spreadsheet.new; end
  def worksheet; @worksheet ||= spreadsheet.s_innapropriate_sheet_name_; end
  def test_c4; assert_in_epsilon(1,worksheet.c4); end
end
end
