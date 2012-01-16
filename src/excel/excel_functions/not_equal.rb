module ExcelFunctions
  
  def not_equal?(a,b)
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)
    
    case a
    when String
      a.downcase != b.downcase
    else
      a != b
    end
  end
  
end
