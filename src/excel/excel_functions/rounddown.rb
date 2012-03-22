module ExcelFunctions
  
  def rounddown(number,decimal_places)
    number = number_argument(number)
    decimal_places = number_argument(decimal_places)
    
    return number if number.is_a?(Symbol)
    return decimal_places if decimal_places.is_a?(Symbol)
    
    (number * 10**decimal_places).truncate.to_f / 10**decimal_places
  end

  
end
