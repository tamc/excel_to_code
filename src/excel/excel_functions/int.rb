module ExcelFunctions
  
  def int(a)
    return 0 if a.nil?
    return :error unless a.is_a? Numeric
    a ||= 0
    result = a.to_s.split('.')[0].to_i
    return result
  end
  
end
