# Tables

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Tables < Spreadsheet
  def b2; "ColA"; end
  def c2; "ColB"; end
  def d2; "Column1"; end
  def b3; 1; end
  def c3; "A"; end
  def d3; string_join(tables.b3,tables.c3); end
  def b4; 2; end
  def c4; "B"; end
  def d4; string_join(tables.b4,tables.c4); end
  def f4; tables.c4; end
  def g4; excel_match("2B",[[tables.b4,tables.c4,tables.d4]],false); end
  def h4; excel_match("B",[[tables.c4,tables.d4]]); end
  def b5; sum([[tables.b3],[tables.b4]]); end
  def c5; sum([[tables.c3],[tables.c4]]); end
  def e6; tables.b2; end
  def f6; tables.c2; end
  def g6; tables.d2; end
  def e7; tables.b5; end
  def f7; tables.c5; end
  def g7; nil; end
  def e8; tables.b2; end
  def f8; tables.c2; end
  def g8; tables.d2; end
  def e9; tables.b3; end
  def f9; tables.c3; end
  def g9; tables.d3; end
  def c10; sum([[tables.b5,tables.c5,nil]]); end
  def e10; tables.b4; end
  def f10; tables.c4; end
  def g10; tables.d4; end
  def c11; sum([[tables.b3],[tables.b4]]); end
  def e11; tables.b5; end
  def f11; tables.c5; end
  def g11; nil; end
  def c12; tables.b5; end
  def c13; sum([[tables.b3,tables.c3,tables.d3],[tables.b4,tables.c4,tables.d4]]); end
  def c14; sum([[tables.b3,tables.c3,tables.d3],[tables.b4,tables.c4,tables.d4]]); end

end
end
