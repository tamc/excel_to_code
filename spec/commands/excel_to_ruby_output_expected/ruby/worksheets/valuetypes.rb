# coding: utf-8
# ValueTypes

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Valuetypes < Spreadsheet
  def a1; @a1 ||= true; end
  def a2; @a2 ||= "Hello"; end
  def a3; @a3 ||= 1; end
  def a4; @a4 ||= 3.1415; end
  def a5; @a5 ||= :name; end
  def a6; @a6 ||= "Hello"; end

end
end
