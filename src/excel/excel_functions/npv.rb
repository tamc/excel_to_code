module ExcelFunctions
  
  def npv(rate, *values)
    # Turn the arguments into numbers
    rate = number_argument(rate)
    values = values.flatten.map { |v| number_argument(v) }

    # Check for errors
    return rate if rate.is_a?(Symbol)
    return :div0 if rate == -1
    values.each { |v| return v if v.is_a?(Symbol) }

    npv = 0

    values.each.with_index { |v, i| npv = npv + (v/((1+rate)**(i+1))) }

    npv
  end
  
end
