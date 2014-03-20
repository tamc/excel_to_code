require_relative 'apply_to_range'

module ExcelFunctions
  
  def less_than_or_equal?(a,b)
    opposite = more_than?(a,b)
    return opposite if opposite.is_a?(Symbol)
    !opposite
  end
  
end
