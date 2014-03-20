require_relative 'apply_to_range'

module ExcelFunctions
  
  def more_than?(a,b)
    # return apply_to_range(a,b) { |a,b| more_than?(a,b) } if a.is_a?(Array) || b.is_a?(Array)
    
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)
    
    a = 0 if a == nil
    b = 0 if b == nil

    case a
    when String
      case b
      when String
        a.downcase > b.downcase
      when Numeric
        true
      when TrueClass, FalseClass
        false
      end
    when TrueClass
      !b.is_a?(TrueClass)
    when FalseClass
      case b
      when TrueClass, FalseClass
        false
      else
        true
      end
    when Numeric
      case b
      when Numeric
        a > b
      else 
        false
      end
    end
  end
  
end
