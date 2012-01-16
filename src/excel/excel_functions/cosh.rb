module ExcelFunctions
  
  def cosh(x)
    return x if x.is_a?(Symbol)
    x ||= 0
    Math.cosh(x)
  end
  
end
