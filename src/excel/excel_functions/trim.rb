module ExcelFunctions
  
  def trim(text)
    return text unless text.is_a?(String)
    text.strip.gsub(/ +/,' ')
  end
  
end
