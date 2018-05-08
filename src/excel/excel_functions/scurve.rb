module ExcelFunctions
  
  def scurve(currentYear, startValue, endValue, duration, startYear = 2018.0)
    currentYear = number_argument(currentYear)
    startValue = number_argument(startValue)
    endValue = number_argument(endValue)
    duration = number_argument(duration)
    startYear = number_argument(startYear)

    return currentYear if currentYear.is_a?(Symbol)
    return startValue if startValue.is_a?(Symbol)
    return endValue if endValue.is_a?(Symbol)
    return duration if duration.is_a?(Symbol)
    return startYear if startYear.is_a?(Symbol)
    
    return startValue if currentYear < startYear 

    x = (currentYear - startYear) / duration.to_f
    x0 = 0.0
    a = endValue - startValue
    sc = 0.999
    eps = 1.0 - sc
    mu = 0.5
    beta = (mu - 1.0) / Math.log(1.0 / sc - 1)
    scurve = a * (((Math.exp(-(x - mu) / beta) + 1) ** -1) - ((Math.exp(-(x0 - mu) / beta) + 1) ** -1)) + startValue
  end

  def halfscurve(currentYear, startValue, endValue , duration, startYear = 2018)
    currentYear = number_argument(currentYear)
    startValue = number_argument(startValue)
    endValue = number_argument(endValue)
    duration = number_argument(duration)
    startYear = number_argument(startYear)

    return currentYear if currentYear.is_a?(Symbol)
    return startValue if startValue.is_a?(Symbol)
    return endValue if endValue.is_a?(Symbol)
    return duration if duration.is_a?(Symbol)
    return startYear if startYear.is_a?(Symbol)
    return startValue if currentYear < startYear 

    scurve(currentYear + duration, startValue, endValue, duration * 2, startYear) - ((endValue - startValue) / 2.0)
  end

  def lcurve(currentYear, startValue, endValue , duration, startYear = 2018)
    currentYear = number_argument(currentYear)
    startValue = number_argument(startValue)
    endValue = number_argument(endValue)
    duration = number_argument(duration)
    startYear = number_argument(startYear)

    return currentYear if currentYear.is_a?(Symbol)
    return startValue if startValue.is_a?(Symbol)
    return endValue if endValue.is_a?(Symbol)
    return duration if duration.is_a?(Symbol)
    return startYear if startYear.is_a?(Symbol)

    return startValue if currentYear < startYear 
    return endValue if currentYear > (startYear + duration)
    return startValue if currentYear < startYear 
    startValue + (endValue - startValue) / duration * (currentYear - startYear)
  end

  def curve(curveType, currentYear, startValue, endValue, duration, startYear = 2018)
    if curveType == "s"
      return scurve(currentYear, startValue, endValue, duration, startYear)
    elsif curveType == "hs"
      return halfscurve(currentYear, startValue, endValue, duration, startYear)
    else
      return lcurve(currentYear, startValue, endValue, duration, startYear)
    end
  end
end
