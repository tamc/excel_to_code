require 'date'
require 'active_support'
require 'active_support/core_ext'

module ExcelFunctions
  
  def month(a)
    return 0 if a.nil?
    return :error unless a.is_a? Numeric
    a ||= 0
    
    result = (Date.civil(1899, 12, 31) + (a - 1).days).month
    return result
  end
  
end
