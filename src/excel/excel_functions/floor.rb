module ExcelFunctions
  
  def floor(number, multiple)
    return number if number.is_a?(Symbol)
    return multiple if multiple.is_a?(Symbol)
    number ||= 0
    return :value unless number.is_a?(Numeric)
    return :value unless multiple.is_a?(Numeric)
    return :div0 if multiple == 0
    return :num if multiple < 0
    number - (number % multiple)
  end
  
end
