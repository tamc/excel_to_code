module ExcelFunctions
  
  def replace(string, start, length, replacement)
    return string if string.is_a?(Symbol)
    return start if start.is_a?(Symbol)
    return length if length.is_a?(Symbol)
    return replacement if replacement.is_a?(Symbol)
    string = string.to_s
    return string + replacement if start >= string.length
    string_join(left(string,start-1),replacement,right(string,string.length - (start - 1 + length)))
  end
  
end
