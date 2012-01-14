# ValueTypes

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Valuetypes < Spreadsheet
  def a1; true; end
  def a5; :name; end
  def a6; "Hello"; end
end
end
