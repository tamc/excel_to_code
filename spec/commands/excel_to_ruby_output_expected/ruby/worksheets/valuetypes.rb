# ValueTypes

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Valuetypes < Spreadsheet
  def a1; true; end
  def a2; "Hello"; end
  def a3; 1; end
  def a4; 3.1415; end
  def a5; :name; end
  def a6; "Hello"; end

end
end
