# FormulaeTypes

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Formulaetypes < Spreadsheet
  def a1; "Simple"; end
  def b1; 2; end
  def a2; "Sharing"; end
  def b2; cosh(multiply(2,pi())); end
  def a3; "Shared"; end
  def b3; cosh(multiply(2,pi())); end
  def a4; "Shared"; end
  def b4; cosh(multiply(2,pi())); end
  def a5; "Array (single)"; end
  def b5; b1; end
  def a6; "Arraying (multiple)"; end
  def b6; excel_if(excel_equal?(b3,8),"Eight","Not Eight"); end
  def a7; "Arrayed (multiple)"; end
  def b7; excel_if(excel_equal?(b4,8),"Eight","Not Eight"); end
  def a8; "Arrayed (multiple)"; end
  def b8; excel_if(excel_equal?(b5,8),"Eight","Not Eight"); end

end
end
