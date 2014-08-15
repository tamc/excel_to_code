module ExcelFunctions
  
  def less_than?(a,b)
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)

    a = 0 if a == nil
    b = 0 if b == nil
    
    case a
    when String
      case b
      when String
        a.downcase < b.downcase
      when Numeric
        false
      when TrueClass, FalseClass
        true
      end
    when TrueClass
      false
    when FalseClass
      b.is_a?(TrueClass)
    when Numeric
      case b
      when Numeric
        a < b
      when String
        true
      when TrueClass, FalseClass
        true
      end
    end
  end
  
end
