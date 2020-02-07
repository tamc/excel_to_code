module ExcelFunctions
  
  def exact(a, b)
    return a if a.is_a?(Symbol)
    return b if b.is_a?(Symbol)

    a = string_argument(a)
    b = string_argument(b)

    return a == b
  end
  
end
