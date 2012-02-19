# Compiled version of /Users/tamc/Documents/github/excel2code/spec/test_data/ExampleSpreadsheet.xlsx
require '/Users/tamc/Documents/github/excel2code/src/excel/excel_functions'

module ExampleSpreadsheet
class Spreadsheet
  include ExcelFunctions
  def valuetypes; @valuetypes ||= Valuetypes.new; end
  def formulaetypes; @formulaetypes ||= Formulaetypes.new; end
  def ranges; @ranges ||= Ranges.new; end
  def referencing; @referencing ||= Referencing.new; end
  def tables; @tables ||= Tables.new; end
  def s_innapropriate_sheet_name_; @s_innapropriate_sheet_name_ ||= S_innapropriate_sheet_name_.new; end
  def common0; @common0 ||= add(referencing.a4,1); end
end
Dir[File.join(File.dirname(__FILE__),"worksheets/","*.rb")].each {|f| autoload(File.basename(f,".rb").capitalize,f)}
end
