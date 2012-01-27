module ExcelFunctions
  
  def subtotal(type,*args)
    case number_argument(type)
    when Symbol; type
    when 1.0, 101.0; average(*args)
    when 2.0, 102.0; count(*args)
    when 3.0, 103.0; counta(*args)
    when 9.0, 109.0; sum(*args)
    else :value
    end
  end  
end
