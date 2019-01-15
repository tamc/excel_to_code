module ExcelFunctions
  
  def countif(count_range, criteria)
    return 0 if criteria.kind_of?(Array) # Who knows why, Excel
    count_range = [count_range] unless count_range.kind_of?(Array)
    rows = count_range.size
    if count_range.first && count_range.first.kind_of?(Array)
        sum_range = Array.new(rows, Array.new(count_range.first.size,1))
    else
        sum_range = Array.new(rows, 1)
    end
    return sumif(count_range, criteria, sum_range)
  end
  
end
