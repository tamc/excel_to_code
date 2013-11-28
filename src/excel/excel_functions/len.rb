module ExcelFunctions
  
  def len(a)
    return a if a.is_a?(Symbol)
    a.to_s.length
  end
  
end
