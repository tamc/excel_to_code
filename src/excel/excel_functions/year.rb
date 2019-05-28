require 'date'

module ExcelFunctions
  
  def year(num)
    year = 1900
    daysInYear = 366;
    
    day = num
    while day > daysInYear
      day -= daysInYear
      year += 1
      daysInYear = year % 4 == 0 ? 366 : 365
    end

    return year
  end
  
end
