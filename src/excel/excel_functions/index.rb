module ExcelFunctions
  
  def index(array, row_number, column_number = :not_specified)
    
    return array if array.is_a?(Symbol)
    return row_number if row_number.is_a?(Symbol)
    return column_number if column_number.is_a?(Symbol) && column_number != :not_specified

    array = [[array]] unless array.is_a?(Array)
        
    if column_number == :not_specified
      if array.length == 1 # It is a single row
        index_for_row_column(array,1,row_number)
      elsif array.first.length == 1 # it is a single column
        index_for_row_column(array,row_number,1)
      else
        return :ref
      end
    else
      if row_number == nil || row_number == 0
        index_for_whole_column(array,column_number)    
      elsif column_number == nil || column_number == 0
        index_for_whole_row(array,row_number)
      else 
        index_for_row_column(array,row_number,column_number)
      end
    end
  end
  
  def index_for_row_column(array,row_number,column_number)
    return :ref if row_number < 1 || row_number > array.length
    row = array[row_number-1]
    return :ref if column_number < 1 || column_number > row.length
    row[column_number-1] || 0
  end
  
  def index_for_whole_row(array,row_number)
    return :ref if row_number < 1
    return :ref if row_number > array.length
    return index_for_row_column(array, row_number, 1) if array.first.length == 1
    [array[row_number-1]]
  end
  
  def index_for_whole_column(array,column_number)
    return :ref if column_number < 1
    return :ref if column_number > array[0].length
    return index_for_row_column(array, 1, column_number) if array.length == 1
    array.map { |row| [row[column_number-1]]}
  end
end
