# coding: utf-8
# FormulaeTypes

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Formulaetypes < Spreadsheet
  def a1; @a1 ||= "Simple"; end
  def b1; @b1 ||= 2; end
  def a2; @a2 ||= "Sharing"; end
  def b2; @b2 ||= 267.7467614837482; end
  def a3; @a3 ||= "Shared"; end
  def b3; @b3 ||= 267.7467614837482; end
  def a4; @a4 ||= "Shared"; end
  def b4; @b4 ||= 267.7467614837482; end
  def a5; @a5 ||= "Array (single)"; end
  def b5; @b5 ||= 2; end
  def a6; @a6 ||= "Arraying (multiple)"; end
  def b6; @b6 ||= "Not Eight"; end
  def a7; @a7 ||= "Arrayed (multiple)"; end
  def b7; @b7 ||= "Not Eight"; end
  def a8; @a8 ||= "Arrayed (multiple)"; end
  def b8; @b8 ||= "Not Eight"; end

end
end
