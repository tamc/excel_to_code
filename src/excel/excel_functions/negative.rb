module ExcelFunctions
  
  def negative(a)
    a = number_argument(a)
  
    return a if a.is_a?(Symbol)
    
    -a

  end
  
end
