module ExcelFunctions
  
  def project_in_array(variable, years, endValue, duration, curveType = "S", startYear = 0, extrapolateCurveType = "LS", relativeEndValue = false)
    project_in_array(1, 1, variable, years, endValue, duration, curveType, startYear, extrapolateCurveType, relativeEndValue)
  end
  
end
