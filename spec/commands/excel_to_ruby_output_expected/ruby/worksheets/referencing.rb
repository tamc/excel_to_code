# coding: utf-8
# Referencing

require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class Referencing < Spreadsheet
  attr_accessor :a4 # Default: 10
  def b4; @b4 ||= add(a4,1); end
  def c4; @c4 ||= add(b4,1); end

  def initialize
    @a4 = 10
  end

end
end
