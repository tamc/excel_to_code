module ExcelFunctions
  
  def isnumber(a)
    return true if a.is_a?(Numeric)
    false
  end
  
end
