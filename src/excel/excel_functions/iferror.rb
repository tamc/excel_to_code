module ExcelFunctions
  
  def iferror(value,value_if_error)
    value_if_error ||= 0
    return value_if_error if value.is_a?(Symbol)
    return value_if_error if value.is_a?(Float) && value.nan?
    value
  end
  
end
