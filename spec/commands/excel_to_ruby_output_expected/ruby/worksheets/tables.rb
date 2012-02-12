# coding: utf-8
# Tables

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Tables < Spreadsheet
  def a1; @a1 ||= add(add(referencing.a4,1),1); end

end
end
