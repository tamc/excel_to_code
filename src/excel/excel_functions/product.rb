module ExcelFunctions
  
  def product(*args)
    args = args.flatten
    args.find {|a| a.is_a?(Symbol)} || args.delete_if { |a| !a.is_a?(Numeric) }.inject(0) { |m,i| m + i }
  end
  
end
