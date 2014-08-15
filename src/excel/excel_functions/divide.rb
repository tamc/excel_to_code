require_relative 'number_argument'

module ExcelFunctions
  
  def divide(a,b)
    a = number_argument(a)
    b = number_argument(b)

    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)
    
    return :div0 if b == 0
    
    a / b.to_f

  rescue ZeroDivisionError
    :div0
  end
  
end
