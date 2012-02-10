require_relative 'apply_to_range'

module ExcelFunctions
  
  def less_than?(a,b)
    # return apply_to_range(a,b) { |a,b| less_than?(a,b) } if a.is_a?(Array) || b.is_a?(Array)
    
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)

    a = 0 if a == nil
    b = 0 if b == nil
    
    case a
    when String
      a.downcase < b.downcase
    when TrueClass, FalseClass
      a = a ? 1 : 0
      b = b ? 1 : 0 if (b.is_a?(TrueClass) || b.is_a?(FalseClass))
      a < b
    else
      a < b
    end
  end
  
end
