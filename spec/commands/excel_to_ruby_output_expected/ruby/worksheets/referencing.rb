# Referencing

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Referencing < Spreadsheet
  def a1; "Named reference"; end
  def a2; referencing.$a$1; end
end
end
