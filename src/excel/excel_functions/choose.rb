module ExcelFunctions
  
  def choose(index,*args)
    # If an argument is an error, return that
    return index if index.is_a?(Symbol)
    error = args.find {|a| a.is_a?(Symbol)}
    return error if error
    
    # If the index is out of bounds, return an error
    return :value unless index
    return :value if index < 1
    return :value if index > args.length
    
    return args[index-1] || 0
    
  end
  
end
