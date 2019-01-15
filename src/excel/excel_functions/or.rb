module ExcelFunctions
  
  def excel_or(*args)
    # Flatten arrays
    args = args.flatten
    
    # If an argument is an error, return that
    error = args.find {|a| a.is_a?(Symbol)}
    return error if error
    
    # Replace 1 and 0 with true and false
    args.map! do |a|
      case a
      when 1; true
      when 0; false
      else; a
      end
    end
    
    # Remove anything not boolean
    args.delete_if { |a| !(a.is_a?(TrueClass) || a.is_a?(FalseClass)) }
    
    # Return an error if nothing less
    return :value if args.empty?
    
    # Now calculate and return
    args.any? {|a| a == true }
 end
  
end
