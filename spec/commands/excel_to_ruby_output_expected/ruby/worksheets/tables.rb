# Tables

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Tables < Spreadsheet
  def b2; "ColA"; end
  def c2; "ColB"; end
  def d2; "Column1"; end
  def g2; [[tables.c2,tables.d2]]; end
  def b3; 1; end
  def c3; "A"; end
  def d3; [[tables.b3,tables.c3]]; end
  def b4; 2; end
  def c4; "B"; end
  def d4; [[tables.b4,tables.c4]]; end
  def f4; tables.c4; end
  def g4; [[tables.b4,tables.c4,tables.d4]]; end
  def h4; [[tables.c4,tables.d4]]; end
  def b5; sum([[tables.b3],[tables.b4]]); end
  def c5; sum([[tables.c3],[tables.c4]]); end
  def e6; [[tables.b2,tables.c2,tables.d2]]; end
  def e7; [[tables.b5,tables.c5,tables.d5]]; end
  def e8; [[tables.b2,tables.c2,tables.d2],[tables.b3,tables.c3,tables.d3],[tables.b4,tables.c4,tables.d4],[tables.b5,tables.c5,tables.d5]]; end
  def e9; sum([[tables.b3,tables.c3,tables.d3],[tables.b4,tables.c4,tables.d4]]); end
  def c10; [[tables.b5,tables.c5,tables.d5]]; end
  def e10; sum([[tables.b2,tables.c2,tables.d2],[tables.b3,tables.c3,tables.d3],[tables.b4,tables.c4,tables.d4],[tables.b5,tables.c5,tables.d5]]); end
  def c11; sum([[tables.b3],[tables.b4]]); end
  def c12; tables.b5; end
end
end
