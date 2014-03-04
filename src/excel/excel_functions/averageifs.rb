module ExcelFunctions
  
  # See sumifs.rb for _filtered_range

  def averageifs(average_range, *criteria)
    filtered = _filtered_range(average_range, *criteria)
    average(*filtered)
  end
  
end
