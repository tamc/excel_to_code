# coding: utf-8
# Tables

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Tables < Spreadsheet
  def a1; @a1 ||= referencing.c4; end
  def b2; @b2 ||= "ColA"; end
  def c2; @c2 ||= "ColB"; end
  def d2; @d2 ||= "Column1"; end
  def b3; @b3 ||= 1; end
  def c3; @c3 ||= "A"; end
  def d3; @d3 ||= "1A"; end
  def b4; @b4 ||= 2; end
  def c4; @c4 ||= "B"; end
  def d4; @d4 ||= "2B"; end
  def f4; @f4 ||= "B"; end
  def g4; @g4 ||= 3; end
  def h4; @h4 ||= 1; end
  def b5; @b5 ||= 3; end
  def c5; @c5 ||= 0; end
  def e6; @e6 ||= "ColA"; end
  def f6; @f6 ||= "ColB"; end
  def g6; @g6 ||= "Column1"; end
  def e7; @e7 ||= 3; end
  def f7; @f7 ||= 0; end
  def g7; @g7 ||= nil; end
  def e8; @e8 ||= "ColA"; end
  def f8; @f8 ||= "ColB"; end
  def g8; @g8 ||= "Column1"; end
  def e9; @e9 ||= 1; end
  def f9; @f9 ||= "A"; end
  def g9; @g9 ||= "1A"; end
  def c10; @c10 ||= 3; end
  def e10; @e10 ||= 2; end
  def f10; @f10 ||= "B"; end
  def g10; @g10 ||= "2B"; end
  def c11; @c11 ||= 3; end
  def e11; @e11 ||= 3; end
  def f11; @f11 ||= 0; end
  def g11; @g11 ||= nil; end
  def c12; @c12 ||= 3; end
  def c13; @c13 ||= 3; end
  def c14; @c14 ||= 3; end

end
end
