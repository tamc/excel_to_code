module ExcelFunctions
  
  def excel_if(condition,true_value,false_value = false)
    return condition if condition.is_a?(Symbol)
    condition ? true_value : false_value
  end
  
end
