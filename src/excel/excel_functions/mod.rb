module ExcelFunctions
  
  def mod(a,b)
    a = number_argument(a)
    b = number_argument(b)
    
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)
    
    return :div0 if b == 0
    
    a % b
  end

end
