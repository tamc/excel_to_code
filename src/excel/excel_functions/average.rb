module ExcelFunctions
  
  def average(*args)
    args = args.flatten
    error = args.find {|a| a.is_a?(Symbol)}
    return error if error
    args.delete_if { |a| !a.is_a?(Numeric) }
    return :div0 if args.empty?
    args.inject(0.0) { |m,i| m + i.to_f } / args.size.to_f
  end
  
end
