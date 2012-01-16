module ExcelFunctions
  
  def more_than_or_equal?(a,b)
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)
    
    case a
    when String
      a.downcase >= b.downcase
    when TrueClass, FalseClass
      a = a ? 1 : 0
      b = b ? 1 : 0 if (b.is_a?(TrueClass) || b.is_a?(FalseClass))
      a >= b
    else
      a >= b
    end
  end
  
end
