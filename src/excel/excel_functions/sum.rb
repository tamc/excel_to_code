module ExcelFunctions
  
  def sum(*args)
    args.flatten.delete_if { |a| !a.is_a?(Numeric) }.inject(0) { |m,i| m + i }
  end
  
end
