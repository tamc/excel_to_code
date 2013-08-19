module ExcelFunctions
  
  def log(a, b = 10)
    a = number_argument(a)
    b = number_argument(b)
    
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)

    return :num if a <= 0
    return :num if b <= 0

    Math.log(a) / Math.log(b)

  end
  
end
