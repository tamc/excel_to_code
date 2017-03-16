module ExcelFunctions
  
  def excel_not(a)
    return a if a.is_a?(Symbol)
    return :value if a.is_a?(String)
    return :value if a.is_a?(Array)
    a = false if a == nil
    a = false if a == 0
    a = true if a.is_a?(Numeric)
    !a
  end
  
end
