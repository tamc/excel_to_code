module ExcelFunctions
  
  def unicode(a)
    return a if a.is_a?(Symbol)
    return :value if a == nil
    a = string_argument(a)
    return a.ord
  end
  
end
