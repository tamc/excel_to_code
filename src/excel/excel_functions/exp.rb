module ExcelFunctions
  
  def exp(a)
    a = number_argument(a)
    return a if a.is_a?(Symbol)

    Math.exp(a)
  end
  
end
