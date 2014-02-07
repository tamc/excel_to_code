module ExcelFunctions
  
  def excel_if(condition,true_value,false_value = false)
    return condition if condition.is_a?(Symbol)
    return false_value if condition == 0
    condition ? true_value : false_value
  end
  
end
