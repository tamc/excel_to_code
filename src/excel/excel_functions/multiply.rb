require_relative 'number_argument'
require_relative 'apply_to_range'

module ExcelFunctions
  
  def multiply(a,b)
  begin
    # return apply_to_range(a,b) { |a,b| multiply(a,b) } if a.is_a?(Array) || b.is_a?(Array)
    
    a = number_argument(a)
    b = number_argument(b)
    
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)
    
    a * b
    
  rescue Error => e
    print e.backtrace.join('\n')
    raise
  end
  end
  
end
