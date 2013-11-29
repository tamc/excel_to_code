require_relative 'apply_to_range'

module ExcelFunctions
  
  def excel_equal?(a,b)
    return apply_to_range(a,b) { |a,b| excel_equal?(a,b) } if a.is_a?(Array) || b.is_a?(Array)
    
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)
    
    return a.downcase == b.downcase if a.is_a?(String) && b.is_a?(String)
    a == b
    
  end
  
end
