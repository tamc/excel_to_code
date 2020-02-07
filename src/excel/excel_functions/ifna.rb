module ExcelFunctions
  
  def ifna(a, b)
    return b || 0 if a == :na
    return a || 0
  end
  
end
