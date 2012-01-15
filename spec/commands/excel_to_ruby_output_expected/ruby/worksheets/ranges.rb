# Ranges

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Ranges < Spreadsheet
  def b1; "This sheet"; end
  def c1; "Other sheet"; end
  def a2; "Standard"; end
  def b2; sum([[f4],[f5],[f6]]); end
  def c2; sum([[valuetypes.a3],[valuetypes.a4]]); end
  def a3; "Column"; end
  def b3; sum([[nil],[nil],[nil],[f4],[f5],[f6]]); end
  def c3; sum([[valuetypes.a1],[valuetypes.a2],[valuetypes.a3],[valuetypes.a4],[valuetypes.a5],[valuetypes.a6]]); end
  def a4; "Row"; end
  def b4; sum([[nil,nil,nil,nil,e5,f5,g5]]); end
  def c4; sum([[valuetypes.a4]]); end
  def f4; 1; end
  def e5; 1; end
  def f5; 2; end
  def g5; 3; end
  def f6; 3; end
end
end
