module ExcelFunctions
  
  def text(number, format)
    number ||= 0
    return "" unless format

    if number.is_a?(String)
      begin
        number = Float(number)
      rescue ArgumentError => e
        # Ignore
      end
    end
    return number unless number.is_a?(Numeric)
    return format if format.is_a?(Symbol)

    case format
    when '0%'
      "#{(number * 100).round}%"
    else
      format
    end
  end
  
end
