module ExcelFunctions
  
  def ln(a)
    a = number_argument(a)
    
    return a if a.is_a?(Symbol)

    return :num if a <= 0

    Math.log(a)

  end
  
end
