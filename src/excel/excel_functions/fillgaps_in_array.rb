module ExcelFunctions
  
  def fillgaps_in_array(columns, rows, variable, years, endYear, extrapolateCurveType = "LS")
    return variable if variable.is_a?(Symbol)
    return years if years.is_a?(Symbol)
    return endYear if endYear.is_a?(Symbol)

    return :ref if columns == 1 && rows == 1
    return :ref unless columns == 1 || rows == 1
    size = max(columns, rows)

    variable = [[variable]] unless variable.is_a?(Array)
    years = [[years]] unless years.is_a?(Array)

    variable.flatten!
    years.flatten!

    return :ref if size > years.count

    result = Array.new(size)

    return :value if years.any? { |y| y.is_a?(Numeric) == false }

    lastYearIndex = years.find_index { |y| y == endYear }
    return :value unless lastYearIndex

    nextIndexes = nextNonBlankIndexes(variable)
    previousIndexes = previousNonBlankIndexes(variable)
    linEst = nil

    i = 0
    while i <= lastYearIndex
      if variable[i] != nil 
        result[i] = variable[i]
      else
        if nextIndexes[i] == nil
          linEst ||= linearEquation(variable.dup, years.dup, extrapolateCurveType)
          result[i] = linEst if linEst.is_a?(Symbol)
          linEst = linEst.first
          p(linEst)
          result[i] = linEst.last + (years[i] * linEst.first)

          # Linear equation
        elsif previousIndexes[i] == nil
          # If empty at start, use next full value
          result[i] = variable[nextIndexes[i]]
        else
          # Linear interpolation
          previous_index = previousIndexes[i]
          next_index = nextIndexes[i]
          previous_value = variable[previous_index]
          next_value = variable[next_index]
          index_difference = next_index - previous_index
          value_difference = next_value - previous_value
          gradient = value_difference.to_f / index_difference.to_f
          relative_position = i - previous_index
          result[i] = previous_value + (relative_position * gradient)
        end
      end
      i += 1
    end

    if rows > columns 
      return result.map { |i| [i] }
    else
      return [result]
    end
  end

  def nextNonBlankIndexes(array)
    result = Array.new(array.size)
    current_index = array.size - 1
    non_blank_index = nil
    while current_index >= 0
      if array[current_index] == nil
        result[current_index] = non_blank_index
      else
        result[current_index] = current_index
        non_blank_index = current_index
      end
      current_index -= 1
    end
    result
  end

  def previousNonBlankIndexes(array)
    result = Array.new(array.size)
    current_index = 0
    non_blank_index = nil
    while current_index < array.size
      if array[current_index] == nil
        result[current_index] = non_blank_index
      else
        result[current_index] = current_index
        non_blank_index = current_index
      end
      current_index += 1
    end
    result
  end

  def linearEquation(yarray, xarray, type)

    # Remove anything that isn't a known value and numeric
    0.upto(xarray.length-1).each do |i|
      xarray[i] = nil unless yarray[i].is_a?(Numeric)
      yarray[i] = nil unless xarray[i].is_a?(Numeric)
    end
    xarray.compact!
    yarray.compact!

    case type
    when "LR"
      return linest(yarray, xarray)

    when "LS"
      slope = slope(yarray, xarray)
      return slope if slope.is_a?(Symbol)
      intercept = yarray.last - (slope * xarray.last)
      return [[slope, intercept]]

    when "L2"
      slope = (yarray[-1] - yarray[-2])/(xarray[-1] - xarray[-2]).to_f
      intercept = yarray.last - (slope * xarray.last)
      return [[slope, intercept]]

    when "F"
      slope = 0
      intercept = yarray.last

    else 
      throw Exception.new("linearEquation type #{type} not recognised")
    end

  end
  
end
