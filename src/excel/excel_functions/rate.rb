module ExcelFunctions
  
  def rate(periods, payment, presentValue, finalValue)
    raise NotSupportedException.new("rate() function has not been implemented fully. Edit src/excel/excel_functions/rate.rb") unless (payment || 0) == 0
    return periods if periods.is_a?(Symbol)
    return presentValue if presentValue.is_a?(Symbol)
    return finalValue if finalValue.is_a?(Symbol)

    return :num unless periods.is_a?(Numeric)
    return :num unless presentValue.is_a?(Numeric)
    return :num unless finalValue.is_a?(Numeric)

    ((finalValue.to_f/(-presentValue.to_f)) ** (1.0/periods))-1.0
  end
  
end
