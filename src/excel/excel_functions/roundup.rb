module ExcelFunctions
  
  def roundup(number,decimal_places)
    number = number_argument(number)
    decimal_places = number_argument(decimal_places)
    
    return number if number.is_a?(Symbol)
    return decimal_places if decimal_places.is_a?(Symbol)
    
    if number < 0
      (number * 10**decimal_places).floor.to_f / 10**decimal_places
    else
      (number * 10**decimal_places).ceil.to_f / 10**decimal_places
    end
  end
  
end
