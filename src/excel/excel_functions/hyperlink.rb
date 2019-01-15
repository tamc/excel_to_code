module ExcelFunctions
  
  def hyperlink(url, text = url)
    u = string_argument(url)
    t = string_argument(text)
    "<a href=#{u.inspect}>#{t}</a>"
  end
  
end
