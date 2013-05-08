module ExcelFunctions
  
  def or(*args)

    args.each do |argument|
      return true if argument
    end

    return false
  end
  
end
