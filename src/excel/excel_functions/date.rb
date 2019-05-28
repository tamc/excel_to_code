require 'date'

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
  
  def validateInput(y, m, d)
    return :num if y < 1900
    return :num if y > 9999
  
    return :num if m < 1
    return :num if m > 12
    
    return :num if d < 1
    daysInMonth = Date.new(y, m, -1, Date::JULIAN).day
    return :num if d > daysInMonth
  end
  
  def date(y, m, d)
    return :num if validateInput(y, m, d) == :num
    
    seq = 0
    year = 1900
    daysInYear = 366
    while y > year
      seq += daysInYear
      year += 1
      daysInYear = year % 4 == 0 ? 366 : 365
    end

    month = 1
    daysInMonth = 31
    while m > month
      seq += daysInMonth
      month += 1
      daysInMonth = Date.new(year, month, -1, Date::JULIAN).day
    end

    return seq + d
  end
  
  
end
