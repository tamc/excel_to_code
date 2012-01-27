module ExcelFunctions
  
  def round(number,decimal_places)
    number = number_argument(number)
    decimal_places = number_argument(decimal_places)
    
    return number if number.is_a?(Symbol)
    return decimal_places if decimal_places.is_a?(Symbol)
    
    number.round(decimal_places)
  end
  
end
