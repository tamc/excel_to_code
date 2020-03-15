module ExcelFunctions
  
  def fillgaps(variable, years, endYear, extrapolateCurveType = "LS")
    fillgaps_in_array(1, 1, variable, years, endYear, extrapolateCurveType)
  end
  
end
