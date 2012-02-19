# coding: utf-8
# Referencing

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Referencing < Spreadsheet
  def a1; @a1 ||= c4; end
  def a2; @a2 ||= c4; end
  attr_accessor :a4 # Default: 10
  def b4; @b4 ||= common0; end
  def c4; @c4 ||= add(common0,1); end
  def a5; @a5 ||= 3; end
  def b8; @b8 ||= c4; end
  def b9; @b9 ||= 3; end
  def b11; @b11 ||= "Named"; end
  def c11; @c11 ||= "Reference"; end
  def d11; @d11 ||= c4; end

  def initialize
    @a4 = 10
  end

end
end
