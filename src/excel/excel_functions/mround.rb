module ExcelFunctions
  
  def mround(value, multiple)
    value = number_argument(value)
    multiple = number_argument(multiple)
    return value if value.is_a?(Symbol)
    return multiple if multiple.is_a?(Symbol)

    # Must both have the same sign
    return :num unless (value < 0) == (multiple < 0)

    # Zeros just return zero
    return 0 if value == 0
    return 0 if multiple == 0

    (value.to_f / multiple.to_f).round * multiple.to_f
  end
  
end
