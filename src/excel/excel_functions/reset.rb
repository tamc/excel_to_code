module ExcelFunctions
  
  # This is a support function for reseting a spreadsheet's instance
  # variables back to nil, allowing the results to be recalculated
  def reset()
    # Set all the instance variables to nil
    instance_variables.each do |iv|
      instance_variable_set(iv,nil)
    end
    # Reset the settable variables to their defaults
    initialize
  end
  
end
