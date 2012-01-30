module ExcelFunctions
  
  def count(*args)
    args = args.flatten
    args.delete_if { |a| !a.is_a?(Numeric)}
    args.size
  end
  
end
