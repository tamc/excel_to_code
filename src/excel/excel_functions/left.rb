module ExcelFunctions
  
  def left(string,characters = 1)
    return string if string.is_a?(Symbol)
    return characters if characters.is_a?(Symbol)
    return nil if string == nil || characters == nil
    return :value if characters < 0
    string = "TRUE" if string == true
    string = "FALSE" if string == false
    string.to_s.slice(0,characters)
  end
  
end
