module ExcelFunctions
  
  def lower(string)
    return string if string.is_a?(Symbol)
    case  string
    when nil; ""
    when String; string.downcase
    when Numeric
      if string.round == string
        string.to_i.to_s
      else
        string.to_s
      end
    when true; "true"
    when false; "false"
    else
      string.to_s
    end
  end
  
end
