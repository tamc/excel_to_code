require_relative 'number_argument'
require_relative 'apply_to_range'

module ExcelFunctions
  
  def divide(a,b)
    return apply_to_range(a,b) { |a,b| divide(a,b) } if a.is_a?(Array) || b.is_a?(Array)

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
