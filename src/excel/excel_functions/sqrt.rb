module ExcelFunctions
  
  def sqrt(a)
    return a if a.is_a?(Symbol)
    a ||= 0
    return :value unless a.is_a?(Numeric)
    return :num if a < 0
    Math.sqrt(a)
  end
  
end
