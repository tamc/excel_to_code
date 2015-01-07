module ExcelFunctions
  
  def iserror(a)
    return true if a.is_a?(Symbol)
    false
  end
  
end
