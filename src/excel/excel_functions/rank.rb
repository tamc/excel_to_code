module ExcelFunctions
  
  def rank(number, list, order = 0)
    number = number_argument(number)
    order = number_argument(order)
    list = [[list]] unless list.is_a?(Array)

    return number if number.is_a?(Symbol)
    return order if order.is_a?(Symbol)
    return list if list.is_a?(Symbol)
    
    ranked = 1
    found = false

    list.flatten.each do |cell|
      return cell if cell.is_a?(Symbol)
      next unless cell.is_a?(Numeric)
      found = true if cell == number
      if order == 0
        ranked += 1 if cell > number
      else
        ranked +=1 if cell < number
      end
    end
    return :na unless found
    return ranked
  end
  
end
