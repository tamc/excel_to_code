require_relative 'number_argument'
require_relative 'apply_to_range'

module ExcelFunctions
  
  def power(a,b)
    a = number_argument(a)
    b = number_argument(b)
    
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)

    return 1 if b ==0 # Special case so can do the following negative number check
    return :num if a < 1 && b < 1
    
    a**b
  end
  
end
