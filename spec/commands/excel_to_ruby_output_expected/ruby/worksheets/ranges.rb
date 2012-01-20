# Ranges

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Ranges < Spreadsheet
  def b1; "This sheet"; end
  def c1; "Other sheet"; end
  def a2; "Standard"; end
  def b2; 6; end
  def c2; 4.141500000000001; end
  def a3; "Column"; end
  def b3; 6; end
  def c3; :name; end
  def a4; "Row"; end
  def b4; 6; end
  def c4; 3.1415; end
  def f4; 1; end
  def e5; 1; end
  def f5; 2; end
  def g5; 3; end
  def f6; 3; end

end
end
