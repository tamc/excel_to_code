# Tables

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Tables < Spreadsheet
  def b2; "ColA"; end
  def c2; "ColB"; end
  def d2; "Column1"; end
  def b3; 1; end
  def c3; "A"; end
  def d3; "1A"; end
  def b4; 2; end
  def c4; "B"; end
  def d4; "2B"; end
  def f4; "B"; end
  def g4; excel_match("2B",[[2,"B","2B"]],false); end
  def h4; excel_match("B",[["B","2B"]]); end
  def b5; sum([[1],[2]]); end
  def c5; sum([["A"],["B"]]); end
  def e6; "ColA"; end
  def f6; "ColB"; end
  def g6; "Column1"; end
  def e7; tables.b5; end
  def f7; tables.c5; end
  def g7; nil; end
  def e8; "ColA"; end
  def f8; "ColB"; end
  def g8; "Column1"; end
  def e9; 1; end
  def f9; "A"; end
  def g9; "1A"; end
  def c10; sum([[tables.b5,tables.c5,nil]]); end
  def e10; 2; end
  def f10; "B"; end
  def g10; "2B"; end
  def c11; sum([[1],[2]]); end
  def e11; tables.b5; end
  def f11; tables.c5; end
  def g11; nil; end
  def c12; tables.b5; end
  def c13; sum([[1,"A","1A"],[2,"B","2B"]]); end
  def c14; sum([[1,"A","1A"],[2,"B","2B"]]); end

end
end
