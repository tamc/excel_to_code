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
    when /0\.(0+)/
      sprintf("%.#{$1.length}f", number)
    else
      raise ExcelToCodeException.new("in TEXT function format #{format} not yet supported by excel_to_code")
    end
  end
  
end
