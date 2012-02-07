# Test for (innapropriate) sheet name!
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestS_innapropriate_sheet_name_
 < Test::Unit::TestCase
  def worksheet; S_innapropriate_sheet_name_
.new; end
end
end
