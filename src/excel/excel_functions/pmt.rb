module ExcelFunctions
  
  def pmt(rate,number_of_periods,present_value, final_value = 0, type = 0)
    
    rate = number_argument(rate)
    number_of_periods = number_argument(number_of_periods)
    present_value = number_argument(present_value)
    final_value = number_argument(final_value)
    type = number_argument(type)

    return rate if rate.is_a?(Symbol)
    return number_of_periods if number_of_periods.is_a?(Symbol)
    return present_value if present_value.is_a?(Symbol)
    return final_value if final_value.is_a?(Symbol)
    return type if type.is_a?(Symbol)
    
    return :num if number_of_periods == 0

    raise NotSupportedException.new("pmt() function non zero final value not implemented") unless final_value == 0
    raise NotSupportedException.new("pmt() function non zero type not implemented") unless type == 0
    
    return -(present_value / number_of_periods) if rate == 0
    -present_value*(rate*((1+rate)**number_of_periods))/(((1+rate)**number_of_periods)-1)
  end
  
end
