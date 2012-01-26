module ExcelFunctions
  
  def abs(a)
    case a
    when Numeric; return a.abs
    when Symbol; return a
    when nil; return 0
    when true; return 1
    when false; return 0
    else; return :value
    end
  end
  
end
