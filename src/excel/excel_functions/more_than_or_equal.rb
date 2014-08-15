module ExcelFunctions
  
  def more_than_or_equal?(a,b)
    opposite = less_than?(a,b)
    return opposite if opposite.is_a?(Symbol)
    !opposite
  end
  
end
