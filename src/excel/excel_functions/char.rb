module ExcelFunctions
  
  def char(a)
    a = number_argument(a)
    return a if a.is_a?(Symbol)
    return :value if a <= 0
    return :value if a >= 256
    a = a.floor
    "".force_encoding("Windows-1252") << a
  end
  
end
