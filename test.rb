#!/usr/bin/ruby -w

# YEAR
# If year is between 0 (zero) and 1899 (inclusive), 
# Excel adds that value to 1900 to calculate the year. 
# For example, DATE(108,1,2) returns January 2, 2008 (1900+108).

# If year is between 1900 and 9999 (inclusive), Excel uses that value as the year. 
# For example, DATE(2008,1,2) returns January 2, 2008.

# If year is less than 0 or is 10000 or greater, 
# Excel returns the #NUM! error value.


# this is the DATE function given valid input
def to_excel_seq_number(year, month, day)
  totalYears = year - 1900
  leapYears = totalYears / 4 + 1
  nonLeapYears = totalYears - leapYears
  seq = (leapYears * 366) + (nonLeapYears * 365)
  
  for a in [1..month-1] do
    daysInMonth = Date.new(year, a, -1).day
    seq = seq + daysInMonth
  end
  seq = seq + d
end

# the is the DAY function
def day_from_excel_seq_number(seq)
  totalYears = seq / 365;
  leapYears = (totalYears / 4) + 1
  nonLeapYears = totalYears - leapYears
  
  year = 1900 + totalYears;
  
  currentMonth = 1
  month = 0
  totalDays = seq - ( (leapYears * 366) + (nonLeapYears * 365) )
  remainingDays = seq - totalDays
  daysInMonth = Date.new(year, currentMonth, -1).day
  while remainingDays > daysInMonth
    month += 1
    currentMonth += 1
    daysInMonth = Date.new(year, currentMonth, -1).day
    remainingDays -= daysInMonth
  end
  # the date is month/remainingDays/year
  return remainingDays  
end

# examples:
# DATE(1900, 14, 29) is 3/1/1901
# DATE(1900, 2, 29) is 2/29/1900
# DATE(1901, -10, 29) is 2/29/1900

puts "1200"
puts date(1200,1,1)
#puts "1900"
#date(1900,1,1)
#puts "-1"
#date(-1,1,1)
puts "10000"
puts date(10000,1,1)


