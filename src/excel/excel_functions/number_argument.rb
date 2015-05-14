module ExcelFunctions
  
  # This is a support function for mapping arguments that are numbers to numeric values
  # deals with the fact that in Excel "2.14" == 2.14, TRUE == 1, FALSE == 0 and nil == 0
  def number_argument(a)
    case a
    when Symbol
      return a
    when String
      begin
        return Float(a)
      rescue ArgumentError
        return :value
      end
    when nil
      return 0
    when true
      return 1
    when false
      return 0
    when Numeric
      return a
    when Array
      return number_argument(a[0][0])
    else
      return :value
    end
  end
  
end
