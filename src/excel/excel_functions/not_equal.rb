require_relative 'apply_to_range'

module ExcelFunctions
  
  def not_equal?(a,b)
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)
    return a.downcase != b.downcase if a.is_a?(String) && b.is_a?(String)
    a != b
  end
  
end
