require_relative 'number_argument'

module ExcelFunctions
  
  def multiply(a,b)
    a = number_argument(a)
    b = number_argument(b)
    
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)
    
    a * b
  end
  
end
