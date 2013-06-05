module ExcelFunctions
  
  def large(range, k)
    range = [range] unless range.is_a?(Array)
    range = range.flatten
    error = range.find {|a| a.is_a?(Symbol)}
    error ||= k if k.is_a?(Symbol)
    return error if error
    range.delete_if { |v| !v.is_a?(Numeric) }
    return :value unless k.is_a?(Numeric)
    return :num unless k>0 && k<=range.size
    range.sort!
    range.reverse!
    range[k-1]
  end
  
end
