# Tables

module ExampleSpreadsheet
class Tables
  def b5; sum([[tables.b3],[tables.b4]]); end
  def c10; [[tables.b5,tables.c5,tables.d5]]; end
  def c11; sum([[tables.b3],[tables.b4]]); end
  def c12; tables.b5; end
  def c5; sum([[tables.c3],[tables.c4]]); end
  def d3; [[tables.b3,tables.c3]]; end
  def d4; [[tables.b4,tables.c4]]; end
  def e10; sum([[tables.b2,tables.c2,tables.d2],[tables.b3,tables.c3,tables.d3],[tables.b4,tables.c4,tables.d4],[tables.b5,tables.c5,tables.d5]]); end
  def e6; [[tables.b2,tables.c2,tables.d2]]; end
  def e7; [[tables.b5,tables.c5,tables.d5]]; end
  def e8; [[tables.b2,tables.c2,tables.d2],[tables.b3,tables.c3,tables.d3],[tables.b4,tables.c4,tables.d4],[tables.b5,tables.c5,tables.d5]]; end
  def e9; sum([[tables.b3,tables.c3,tables.d3],[tables.b4,tables.c4,tables.d4]]); end
  def f4; tables.c4; end
  def g2; [[tables.c2,tables.d2]]; end
  def g4; [[tables.b4,tables.c4,tables.d4]]; end
  def h4; [[tables.c4,tables.d4]]; end
end
end
