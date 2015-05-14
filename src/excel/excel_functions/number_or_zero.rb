module ExcelFunctions
  
  def number_or_zero(a)
    return a if a.is_a?(Symbol)
    return a if a.is_a?(Numeric)
    0
  end
  
end
