module ExcelFunctions
  
  def counta(*args)
    args.flatten!
    args.compact!
    args.size
  end
  
end
