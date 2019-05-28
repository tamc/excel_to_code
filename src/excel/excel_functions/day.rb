require 'date'

module ExcelFunctions

  def day(num)
    year = 1900
    daysInYear = 366;
    
    day = num
    while day > daysInYear
      day -= daysInYear
      year += 1
      daysInYear = year % 4 == 0 ? 366 : 365
    end
    
    month = 1
    daysInMonth = Date.new(year, month, -1, Date::JULIAN).day
    
    while day > daysInMonth
      month += 1
      day -= daysInMonth
      daysInMonth = Date.new(year, month, -1, Date::JULIAN).day
    end

    return day
  end
  
end
