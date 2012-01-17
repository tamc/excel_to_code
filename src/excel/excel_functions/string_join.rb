module ExcelFunctions
  
  def string_join(*strings)
    strings.find {|s| s.is_a?(Symbol)} || strings.map { |s| s == nil ? "0" : s.to_s }.join('')
  end
  
end
