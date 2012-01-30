require_relative '../excel'
require_relative '../util'

class Table
  
  attr_accessor :name
  
  def initialize(name,worksheet,reference,number_of_total_rows,*column_name_array)
    @name, @worksheet, @area, @number_of_total_rows, @column_name_array = name, worksheet, Area.for(reference), number_of_total_rows.to_i, column_name_array.map(&:downcase)
  end
  
  def reference_for(table_name,structured_reference,calling_worksheet,calling_cell)
    raise NotSupportedException.new("Local table reference not supported in #{structured_reference.inspect}") unless table_name
    @area.calculate_excel_variables
    case structured_reference
    when /\[#Headers\],\[(.*?)\]:\[(.*?)\]/io
      column_number_start = @column_name_array.find_index($1.downcase)
      column_number_finish = @column_name_array.find_index($2.downcase)
      ast_for_area @area.excel_start.offset(0,column_number_start), @area.excel_start.offset(0,column_number_finish)
    when /\[#Totals\],\[(.*?)\]:\[(.*?)\]/io
      column_number_start = @column_name_array.find_index($1.downcase)
      column_number_finish = @column_name_array.find_index($2.downcase)
      ast_for_area @area.excel_start.offset(@area.height,column_number_start), @area.excel_start.offset(@area.height,column_number_finish)
    when /\[#This Row\],\[(.*?)\]:\[(.*?)\]/io
      r = Reference.for(calling_cell)
      r.calculate_excel_variables
      row = r.excel_row_number
      column_number_start = @column_name_array.find_index($1.downcase)
      column_number_finish = @column_name_array.find_index($2.downcase)
      ast_for_area @area.excel_start.offset(row - @area.excel_start.excel_row_number,column_number_start), @area.excel_start.offset(row - @area.excel_start.excel_row_number,column_number_finish)      
    when /\[#Headers\],\[(.*?)\]/io
      column_number = @column_name_array.find_index($1.downcase)
      ast_for_cell @area.excel_start.offset(0,column_number)
    when /\[#Totals\],\[(.*?)\]/io
      column_number = @column_name_array.find_index($1.downcase)
      ast_for_cell @area.excel_start.offset(@area.height,column_number)
    when /\[#This Row\],\[(.*?)\]/io
      r = Reference.for(calling_cell)
      r.calculate_excel_variables
      row = r.excel_row_number
      column_number = @column_name_array.find_index($1.downcase)
      ast_for_cell @area.excel_start.offset(row - @area.excel_start.excel_row_number,column_number)      
    when /#Headers/io
      ast_for_area @area.excel_start.offset(0,0), @area.excel_start.offset(0,@area.width)
    when /#Totals/io
      ast_for_area @area.excel_start.offset(@area.height,0), @area.excel_start.offset(@area.height,@area.width)
    when /#Data/io, ""
      ast_for_area @area.excel_start.offset(1,0), @area.excel_finish.offset(-@number_of_total_rows,0)
    when /#All/io, ""
      ast_for_area @area.excel_start, @area.excel_finish
    when /#This Row/io
      r = Reference.for(calling_cell)
      r.calculate_excel_variables
      row = r.excel_row_number
      ast_for_area "#{@area.excel_start.excel_column}#{row}", "#{@area.excel_finish.excel_column}#{row}"
    else
      if calling_worksheet == @worksheet && @area.includes?(calling_cell)
        r = Reference.for(calling_cell)
        r.calculate_excel_variables
        row = r.excel_row_number
        column_number = @column_name_array.find_index(structured_reference.downcase)
        ast_for_cell @area.excel_start.offset(row - @area.excel_start.excel_row_number,column_number)      
      else        
        column_number = @column_name_array.find_index(structured_reference.downcase)
        ast_for_area @area.excel_start.offset(1,column_number), @area.excel_start.offset(@area.height - @number_of_total_rows,column_number)
      end
    end
  end
  
  def ast_for_area(start,finish)
    [:sheet_reference,@worksheet,[:area,start,finish]]
  end
  
  def ast_for_cell(ref)
    [:sheet_reference,@worksheet,[:cell,ref]]
  end

end
