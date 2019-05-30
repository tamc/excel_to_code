module ExcelFunctions
  
  def sum(*args)
    args = args.flatten
    result = args.find {|a| a.is_a?(Symbol)} || args.delete_if { |a| !a.is_a?(Numeric) }.inject(0) { |m,i| m + i }

    return :num if result.is_a?(Numeric) && result.infinite?

    result
  end
  
end
