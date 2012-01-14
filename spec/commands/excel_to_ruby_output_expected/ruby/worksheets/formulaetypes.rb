# FormulaeTypes

module ExampleSpreadsheet
class Formulaetypes
  def b1; add(1,1); end
  def b2; cosh(multiply(2,pi())); end
  def b3; cosh(multiply(2,pi())); end
  def b4; cosh(multiply(2,pi())); end
  def b5; b1; end
  def b6; excel_if(equal?(b3,8),"Eight","Not Eight"); end
  def b7; excel_if(equal?(b4,8),"Eight","Not Eight"); end
  def b8; excel_if(equal?(b5,8),"Eight","Not Eight"); end
end
end
