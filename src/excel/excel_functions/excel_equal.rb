module ExcelFunctions
  
  def excel_equal?(a,b)
    case a
    when String
      a.downcase == b.downcase
    else
      a == b
    end
  end
  
end
