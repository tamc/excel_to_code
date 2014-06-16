module ExcelFunctions
  
  def iserr(a)
    return false if a == :na
    return true if a.is_a?(Symbol)
    false
  end
  
end
