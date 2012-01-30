module ExcelFunctions
  
  def sumif(check_range,criteria,sum_range = check_range)
    sumifs(sum_range,check_range,criteria)
  end
  
end
