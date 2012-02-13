# coding: utf-8
# Ranges

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Ranges < Spreadsheet
  def b1; @b1 ||= "This sheet"; end
  def c1; @c1 ||= "Other sheet"; end
  def a2; @a2 ||= "Standard"; end
  def b2; @b2 ||= 6; end
  def c2; @c2 ||= 4.141500000000001; end
  def a3; @a3 ||= "Column"; end
  def b3; @b3 ||= 6; end
  def c3; @c3 ||= :name; end
  def a4; @a4 ||= "Row"; end
  def b4; @b4 ||= 6; end
  def c4; @c4 ||= 3.1415; end
  def f4; @f4 ||= 1; end
  def e5; @e5 ||= 1; end
  def f5; @f5 ||= 2; end
  def g5; @g5 ||= 3; end
  def f6; @f6 ||= 3; end

end
end
