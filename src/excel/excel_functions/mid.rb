module ExcelFunctions
  
  def mid(text, start_num, num_chars)
    start_num = number_argument(start_num)
    num_chars = number_argument(num_chars)

    return text if text.is_a?(Symbol)
    return start_num if start_num.is_a?(Symbol)
    return num_chars if num_chars.is_a?(Symbol)

    text = text.to_s

    return :value if start_num < 1
    return :value if num_chars < 0

    return "" if start_num > text.length
    text[start_num - 1, num_chars]
  end
  
end
