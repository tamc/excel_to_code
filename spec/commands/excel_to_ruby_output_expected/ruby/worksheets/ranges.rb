# Ranges

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Ranges < Spreadsheet
  def b1; "This sheet"; end
  def c1; "Other sheet"; end
  def a2; "Standard"; end
  def b2; sum([[1],[2],[3]]); end
  def c2; sum([[1],[3.1415]]); end
  def a3; "Column"; end
  def b3; sum([[nil],[nil],[nil],[1],[2],[3]]); end
  def c3; sum([[true],["Hello"],[1],[3.1415],[:name],["Hello"]]); end
  def a4; "Row"; end
  def b4; sum([[nil,nil,nil,nil,1,2,3]]); end
  def c4; sum([[3.1415]]); end
  def f4; 1; end
  def e5; 1; end
  def f5; 2; end
  def g5; 3; end
  def f6; 3; end

end
end
