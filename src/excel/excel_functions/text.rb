module ExcelFunctions
  
  def text(number, format)
    number ||= 0
    return "" unless format
    format = "0" if format == 0.0

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
    when /^(0(\.0*)?)%/
      text(number*100, $1)+"%"
    when /^(0+)$/
      sprintf("%0#{$1.length}.0f", number)
    when /#,#+(0\.0+)?/
      formated_with_decimals = text(number, $1 || "0")
      parts = formated_with_decimals.split('.')
      parts[0].gsub!(/(\d)(?=(\d\d\d)+(?!\d))/, "\\1,")
      parts.join('.')
    when /0\.(0+)/
      sprintf("%.#{$1.length}f", number)
    else
      raise ExcelToCodeException.new("in TEXT function format #{format} not yet supported by excel_to_code")
    end
  end
  
end
