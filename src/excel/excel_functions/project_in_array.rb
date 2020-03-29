module ExcelFunctions
  
  def project_in_array(columns, rows, variable, years, endValue, duration, curveType = "S", startYear = 0, extrapolateCurveType = "LS", relativeEndValue = false)
    return variable if variable.is_a?(Symbol)
    return years if years.is_a?(Symbol)
    return endValue if endValue.is_a?(Symbol)
    return duration if duration.is_a?(Symbol)
    return curveType if curveType.is_a?(Symbol)
    return startYear if startYear.is_a?(Symbol)
    return extrapolateCurveType if extrapolateCurveType.is_a?(Symbol)
    return relativeEndValue if relativeEndValue.is_a?(Symbol)

    return :ref unless columns == 1 || rows == 1
    size = max(columns, rows)

    variable = [[variable]] unless variable.is_a?(Array)
    years = [[years]] unless years.is_a?(Array)

    variable.flatten!
    years.flatten!

    return :ref if size > years.count

    result = Array.new(size)

    return :value if years.any? { |y| y.is_a?(Numeric) == false }

    lastYearWithValue = 1
    lastYearWithValueIndex = 1

    (0..variable.length).each do |i|
      if variable[i].is_a?(Numeric) == false 
        lastYearWithValueIndex = i - 1
        lastYearWithValue = years[lastYearWithValueIndex]
        break
      end
    end

    if relativeEndValue 
      endValue = endValue * variable[lastYearWithValueIndex]
    end

    startYearWithIndex = 0
    if startYear == 0 
      startYear = lastYearWithValue
      startYearIndex = lastYearWithValueIndex
    else
      startYearIndex = years.find_index { |y| y == startYear }
      return :value unless startYearIndex
    end

    if startYearIndex < lastYearWithValueIndex 
      (0..result.length).each do |i|
        if i < startYearIndex
          result[i] = variable[i]
        else
          result[i] = curve(curveType, years[i], variables[startYearIndex], endValue, duration, startYear)
        end
      end
    else
      knownValuesArray = variable[0..lastYearWithValueIndex]
      knownYearsArray = years[0..lastYearWithValueIndex]
        
      f = linearEquation(knownValuesArray.dup, knownYearsArray.dup, extrapolateCurveType)

      (0..result.length).each do |i|
        if i <= lastYearWithValueIndex
          result[i] = variable[i]
        elsif i < startYearIndex
          result[i] = (f.first * years[i]) + f.last
        else
          result[i] = curve(curveType, years[i], result[startYearIndex], endValue, duration, startYear)
        end
      end
    end

    if rows > columns 
      return result.map { |i| [i] }
    else
      return [result]
    end
  end
  
end
