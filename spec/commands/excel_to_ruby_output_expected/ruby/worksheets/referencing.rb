# Referencing

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Referencing < Spreadsheet
  def a1; "Named reference"; end
  def a2; referencing.a1; end
end
end
