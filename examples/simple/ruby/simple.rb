# coding: utf-8
# Compiled version of /Users/tamc/Documents/github/excel_to_code/examples/simple/simple.xlsx
require '/usr/local/lib/ruby/gems/2.5.0/gems/excel_to_code-0.3.17/src/excel/excel_functions'

class Simple
  include ExcelFunctions
  def original_excel_filename; "/Users/tamc/Documents/github/excel_to_code/examples/simple/simple.xlsx"; end
  def inputs_a1; @inputs_a1 ||= "Your first name:"; end
  attr_accessor :inputs_b1 # Default: "Harry"
  def inputs_a2; @inputs_a2 ||= "Your surname:"; end
  attr_accessor :inputs_b2 # Default: "Potter"
  def inputs_a4; @inputs_a4 ||= "What is the meaning of life?"; end
  attr_accessor :inputs_b4 # Default: 24.0
  def outputs_a1; @outputs_a1 ||= string_join("Hello ",inputs_b1," ",inputs_b2); end
  def outputs_a2; @outputs_a2 ||= excel_if(excel_equal?(inputs_b4,42.0),"Well done, you know the meaning of life","Oh dear, that isn't the meaning of life. Guess again."); end
  def outputs_a3; @outputs_a3 ||= string_join("You were ",abs(subtract(42.0,inputs_b4))," out."); end

# Start of named references
# End of named references

  # starting initializer
  def initialize
    @inputs_b1 = "Harry"
    @inputs_b2 = "Potter"
    @inputs_b4 = 24.0
  end

end
