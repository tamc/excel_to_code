# Referencing

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Referencing < Spreadsheet
  def a1; "Named reference"; end
  def a2; "Named reference"; end
  attr_accessor :a4 # Default: 10
  def b4; add(a4,1); end
  def c4; add(b4,1); end
  def a5; 3; end
  def b8; "Named reference"; end
  def b9; sum([[tables.b5,tables.c5,nil]]); end
  def b11; "Named"; end
  def c11; "Reference"; end
  def d11; "Named reference"; end

  def initialize
    @a4 = 10
  end

end
end
