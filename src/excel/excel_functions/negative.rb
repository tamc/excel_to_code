module ExcelFunctions
  
  def negative(a)
    return a.map { |c| negative(c) } if a.is_a?(Array)

    a = number_argument(a)
  
    return a if a.is_a?(Symbol)
    
    -a

  end
  
end
