module ExcelFunctions
  
  ## For now we are only going to accept valid dates for input.
  ## TODO: fix up rules to accomodate other values according to Excel's rules
  # According to Excel docs:
  
  # YEAR
  # If year is between 0 (zero) and 1899 (inclusive), 
  # Excel adds that value to 1900 to calculate the year. 
  # For example, DATE(108,1,2) returns January 2, 2008 (1900+108).
  
  # If year is between 1900 and 9999 (inclusive), 
  # Excel uses that value as the year. 
  # For example, DATE(2008,1,2) returns January 2, 2008.
  
  # If year is less than 0 or is 10000 or greater, 
  # Excel returns the #NUM! error value.
  

  # MONTH
  # If month is greater than 12, 
  # month adds that number of months to the first month in the year specified. 
  # For example, DATE(2008,14,2) 
  # returns the serial number representing February 2, 2009.
  
  # If month is less than 1, 
  # month subtracts the magnitude of that number of months, plus 1, 
  # from the first month in the year specified. 
  # For example, DATE(2008,-3,2) returns the serial number 
  # representing September 2, 2007
  
  # DAY
  
  # If day is greater than the number of days in the month specified, 
  # day adds that number of days to the first day in the month. 
  # For example, DATE(2008,1,35) returns the serial number 
  # representing February 4, 2008.
  
  # If day is less than 1, 
  # day subtracts the magnitude that number of days, plus one, 
  # from the first day of the month specified. 
  # For example, DATE(2008,1,-15) returns the serial number 
  # representing December 16, 2007.
  
  def date(y, m, d)
    raise NotSupportedException.new("date() function has not been implemented fully. Edit src/excel/excel_functions/date.rb")
    # return a if a.is_a?(Symbol)
    # a ||= 0
    # implement function
    # return result
    y = number_argument(y)
    m = number_argument(m)
    d = number_argument(d)
    
    return y if y.is_a?(Symbol)
    return m if m.is_a?(Symbol)
    return d if d.is_a?(Symbol)

    return :error if !(y > 1899 && y <= 9999) 
    return :error if !(m > 0 && m <= 12)
    
    daysInMonth = Date.new(year, month, -1).day
    return :error if !(d > 0 && d <= daysInMonth)

    return 
  end
  
end
