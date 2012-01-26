module ExcelFunctions
  
  def max(*args)
    args = args.flatten
    error = args.find {|a| a.is_a?(Symbol)}
    return error if error
    args.delete_if { |a| !a.is_a?(Numeric) }
    return 0 if args.empty?
    args.max
  end
  
end
