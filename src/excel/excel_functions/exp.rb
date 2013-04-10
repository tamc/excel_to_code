module ExcelFunctions
  
  def exp(a)
    return 0 if a.nil?
    return :error unless a.is_a? Numeric
    a ||= 0
    result = Math::E ** a
    return result
  end
  
end
