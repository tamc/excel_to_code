module ExcelFunctions
  
  def mod(a,b)
    case a
    when Symbol
      return a
    when String
      begin
        a = Float(a)
      rescue ArgumentError
        return :value
      end
    when nil
      a = 0
    when true
      a = 1
    when false
      a = 0
    when Numeric
       # ok
    else
      return :value
    end
    
    case b
    when Symbol
      return b
    when String
      begin
        b = Float(b)
      rescue ArgumentError
        return :value
      end
    when nil
      b = 0
    when true
      b = 1
    when false
      b = 0
    when Numeric
      # ok
    else
      return :value
    end
    
    return :div0 if b == 0
    
    a % b
  end
  
end
