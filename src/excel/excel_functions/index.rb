module ExcelFunctions
  
  def index(array,row_number,column_number = nil)
    
    return array if array.is_a?(Symbol)
    return row_number if row_number.is_a?(Symbol)
    return column_number if column_number.is_a?(Symbol)

    array = [[array]] unless array.is_a?(Array)
  
    if column_number == nil
      if array.length == 1 # It is a single row
        row = array.first
        return :ref if row_number > row.length
        return :ref if row_number < 1
        return row[row_number-1] || 0
      elsif array.all? { |row| row.length == 1 } # it is a single column
        return :ref if row_number > array.length
        return :ref if row_number < 1
        return array[row_number - 1][0] || 0
      else
        return :ref
      end
    elsif row_number == nil
      return :ref unless array.length == 1
      return :ref if column_number < 1
      return :ref if column_number > array[0].length
      return array[0][column_number-1] || 0
    else
      if row_number == 0
        return :ref if column_number < 1
        return :ref if column_number > array[0].length
        return array.map { |row| row[column_number-1] }
      elsif column_number == 0
        return :ref if row_number < 1
        return :ref if row_number > array.length
        return array[row_number-1]
      else
        return :ref if row_number < 1 || row_number > array.length
        row = array[row_number-1]
        return :ref if column_number < 1 || column_number > row.length
        row[column_number-1] || 0
      end

    end
  end
  
end
