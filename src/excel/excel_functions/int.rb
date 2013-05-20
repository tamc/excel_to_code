module ExcelFunctions
  
  def int(number)
    number = number_argument(number)
    return number if number.is_a?(Symbol)
    number.floor.to_f
  end

  
end
