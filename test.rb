

require 'date'


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
  
  puts sprintf("year: %d, seq: %d", year, seq)
  
  month = 1
  daysInMonth = 31
  while m > month
    seq += daysInMonth
    month += 1
    daysInMonth = Date.new(year, month, -1, Date::JULIAN).day
  end
  puts sprintf("year: %d, month: %d seq: %d", year, month, seq)
  puts sprintf("returning: %d", seq + d)
  return seq + d
end

# the is the DAY function
def day(seq)
  arr = extract(seq)
  return arr[2] 
end

# this is the MONTH function
def month(seq)
  arr = extract(seq)
  return arr[1]
end

# this is the YEAR function
def year(seq)
  arr = extract(seq)
  return arr[0]
end


def fixup_input(y, m, d)
  
  if y < 0 || y > 9999
    return :error
  end
  
  year = y + 1900 if y >= 0 && y <= 1899
  year = y if y >= 1900 && y <= 9999
  
  month = m
  if m > 12
    month = m % 12
    year += m/12
    :error if year > 9999
  elsif m < 1
    year -= (m.abs/12 + 1)
    if year < 1900
      year += 1900
    end
    month = m % 12
  end
  
  daysInMonth = Date.new(year, month, -1, Date::JULIAN).day
  if d > daysInMonth
    while d > daysInMonth
      month += 1
      if month > 12
        month = 1
        year += 1
        :error if year > 9999
        d -= daysInMonth
        daysInMonth = Date.new(year, month, -1, Date::JULIAN).day
      else
        d -= daysInMonth
        daysInMonth = Date.new(year, month, -1, Date::JULIAN).day
      end
    end
  elsif d < 1
    while d < 0
      month -= 1
      if month < 1
        year -= 1
        if year < 1900
          return :error
        end
        month = 12
      end
      daysInMonth = Date.new(year, month, -1, Date::JULIAN).day
      d += daysInMonth
    end
    if d == 0
      month -= 1
      if month < 1
        year -= 1
        if year < 1900
          return :error
        end
        month = 12
      end
      d = Date.new(year, month, -1, Date::JULIAN).day
    end
  else # d == 0
    # this is the last day of the prior month
    month -= 1
    if month < 1
      year -= 1
      sprintf("YEAR: %d: ", year)
      if year < 1900
        return :error
      end
      month = 12
    end
    d = Date.new(year, month, -1, Date::JULIAN).day
  end
  
  return [year, month, d]
end



def printDate(d) 
  sprintf("%d/%d/%d", d[1], d[2], d[0])
end




def validateInput(y, m, d)
  return :num if y < 1900
  return :num if y > 9999

  return :num if m < 1
  return :num if m > 12
  
  return :num if d < 1
  daysInMonth = Date.new(y, m, -1, Date::JULIAN).day
  return :num if d > daysInMonth
end


def extract(num)

  year = 1900
  daysInYear = 366;
  
  day = num
  while day > daysInYear
    day -= daysInYear
    year += 1
    daysInYear = year % 4 == 0 ? 366 : 365
  end
  
  month = 1
  puts "Year: ", year
  
  daysInMonth = Date.new(year, month, -1, Date::JULIAN).day
  puts "Days in month ", daysInMonth
  
  while day > daysInMonth
    month += 1
    day -= daysInMonth
    daysInMonth = Date.new(year, month, -1, Date::JULIAN).day
    puts printf("There are %d days in the month %d", daysInMonth, month)
  end
  puts printf("year: %d, month %d, day %d", year, month, day)
  return [year, month, day]
end

def runExtract
  File.open("./dateinput.csv") do |file|
    file.each do |line|
      arr = line.split(",")

      puts sprintf("seq: %d, day: %d, month: %d, year: %d", arr[0], arr[1], arr[2], arr[3])
      theDate = extract(arr[0].to_i)
      puts printDate(theDate)
      if theDate[0] != arr[3].to_i || theDate[1] != arr[2].to_i  || theDate[2] != arr[1].to_i
        puts "FAILED, sequence number ", arr[0]
        exit 1
      end
    end
  end
end

def runToDate
  File.open("./dateinput.csv") do |file|
    file.each do |line|
      arr = line.split(",")

      puts sprintf("seq: %d, day: %d, month: %d, year: %d", arr[0], arr[1], arr[2], arr[3])
      
      seq = date(arr[3].to_i, arr[2].to_i, arr[1].to_i)

      puts sprintf("date returned seq %s", seq.to_s)
      if seq != arr[0].to_i
        puts "FAILED, sequence number ", arr[0]
        exit 1
      end
    end
  end
end

runToSeq
#date(1900,1,1)
#runExtract
#extract(24830)

#theDate = extract_from_excel_seq_number(37)
#
#theDate = extract_from_excel_seq_number(24890)