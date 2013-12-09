require_relative '../excel'
require_relative '../util'

class Table
  
  attr_accessor :name
  
  def initialize(name,worksheet,reference,number_of_total_rows,*column_name_array)
    @name, @worksheet, @area, @number_of_total_rows, @column_name_array = name, worksheet.to_sym, Area.for(reference), number_of_total_rows.to_i, column_name_array.map { |c| c.strip.downcase }
    @area.calculate_excel_variables    
    @data_area = Area.for("#{@area.excel_start.offset(1,0)}:#{@area.excel_finish.offset(-@number_of_total_rows,0)}")
    @data_area.calculate_excel_variables
  end
  
  def reference_for(table_name,structured_reference,calling_worksheet,calling_cell)
    raise NotSupportedException.new("Local table reference not supported in #{structured_reference.inspect}") unless table_name

    case structured_reference
    when /\[#Headers\],\[(.*?)\]:\[(.*?)\]/io
      column_number_start = @column_name_array.find_index($1.strip.downcase)
      column_number_finish = @column_name_array.find_index($2.strip.downcase)
      return ref_error unless column_number_start && column_number_finish
      ast_for_area @area.excel_start.offset(0,column_number_start), @area.excel_start.offset(0,column_number_finish)

    when /\[#Totals\],\[(.*?)\]:\[(.*?)\]/io
      column_number_start = @column_name_array.find_index($1.strip.downcase)
      column_number_finish = @column_name_array.find_index($2.strip.downcase)
      return ref_error unless column_number_start && column_number_finish
      ast_for_area @area.excel_start.offset(@area.height,column_number_start), @area.excel_start.offset(@area.height,column_number_finish)

    when /\[#This Row\],\[(.*?)\]:\[(.*?)\]/io
      r = Reference.for(calling_cell)
      r.calculate_excel_variables
      row = r.excel_row_number
      column_number_start = @column_name_array.find_index($1.strip.downcase)
      column_number_finish = @column_name_array.find_index($2.strip.downcase)
      return ref_error unless column_number_start && column_number_finish
      ast_for_area @area.excel_start.offset(row - @area.excel_start.excel_row_number,column_number_start), @area.excel_start.offset(row - @area.excel_start.excel_row_number,column_number_finish)

    when /\[(.*?)\]:\[(.*?)\]/io
      column_number_start = @column_name_array.find_index($1.strip.downcase)
      column_number_finish = @column_name_array.find_index($2.strip.downcase)
      return ref_error unless column_number_start && column_number_finish
      ast_for_area @area.excel_start.offset(1,column_number_start), @area.excel_start.offset(@area.height - @number_of_total_rows,column_number_finish)

    when /\[#Headers\],\[(.*?)\]/io
      column_number = @column_name_array.find_index($1.strip.downcase)
      return ref_error unless column_number      
      ast_for_cell @area.excel_start.offset(0,column_number)

    when /\[#Totals\],\[(.*?)\]/io
      column_number = @column_name_array.find_index($1.strip.downcase)
      return ref_error unless column_number
      ast_for_cell @area.excel_start.offset(@area.height,column_number)

    when /\[#This Row\],\[(.*?)\]/io
      r = Reference.for(calling_cell)
      r.calculate_excel_variables
      row = r.excel_row_number
      column_number = @column_name_array.find_index($1.strip.downcase)
      return ref_error unless column_number
      ast_for_cell @area.excel_start.offset(row - @area.excel_start.excel_row_number,column_number)      

    when /#Headers/io
      if calling_worksheet == @worksheet && @data_area.includes?(calling_cell)
        r = Reference.for(calling_cell)
        r.calculate_excel_variables
        ast_for_cell "#{r.excel_column}#{@area.excel_start.excel_row_number}"
      else
        ast_for_area @area.excel_start.offset(0,0), @area.excel_start.offset(0,@area.width)
      end

    when /#Totals/io
      if calling_worksheet == @worksheet && @data_area.includes?(calling_cell)
        r = Reference.for(calling_cell)
        r.calculate_excel_variables
        ast_for_cell "#{r.excel_column}#{@area.excel_finish.excel_row_number}"
      else
        ast_for_area @area.excel_start.offset(@area.height,0), @area.excel_start.offset(@area.height,@area.width)
      end

    when /#Data/io, ""
      ast_for_area @data_area.excel_start, @data_area.excel_finish

    when /#All/io
      ast_for_area @area.excel_start, @area.excel_finish

    when /#This Row/io
      r = Reference.for(calling_cell)
      r.calculate_excel_variables
      row = r.excel_row_number
      ast_for_area "#{@area.excel_start.excel_column}#{row}", "#{@area.excel_finish.excel_column}#{row}"

    else
      if calling_worksheet == @worksheet && @data_area.includes?(calling_cell)
        r = Reference.for(calling_cell)
        r.calculate_excel_variables
        row = r.excel_row_number
        column_number = @column_name_array.find_index(structured_reference.strip.downcase)
        return ref_error unless column_number
        ast_for_cell @area.excel_start.offset(row - @area.excel_start.excel_row_number,column_number)      
      else        
        column_number = @column_name_array.find_index(structured_reference.strip.downcase)
        return ref_error unless column_number
        ast_for_area @area.excel_start.offset(1,column_number), @area.excel_start.offset(@area.height - @number_of_total_rows,column_number)
      end
    end
  end
  
  def ast_for_area(start,finish)
    [:sheet_reference,@worksheet,[:area,start.to_sym,finish.to_sym]]
  end
  
  def ast_for_cell(ref)
    [:sheet_reference,@worksheet,[:cell,ref.to_sym]]
  end
  
  def ref_error
    [:error,"#REF!"]
  end
  
  def includes?(sheet,reference)
    return false unless @worksheet == sheet
    @area.includes?(reference)
  end
  
end
