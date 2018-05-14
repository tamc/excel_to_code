module ExcelFunctions
  
  def ceiling(number, multiple, mode = 0)
    return number if number.is_a?(Symbol)
    return multiple if multiple.is_a?(Symbol)
    return mode if mode.is_a?(Symbol)
    number = number_argument(number)
    multiple = number_argument(multiple)
    mode = number_argument(mode)
    return :value unless number.is_a?(Numeric)
    return :value unless multiple.is_a?(Numeric)
    return :value unless mode.is_a?(Numeric)
    return 0 if multiple == 0
    if mode == 0 || number > 0
      whole, remainder = number.divmod(multiple)
      num_steps = remainder > 0 ? whole + 1 : whole
      num_steps * multiple
    else # Need to round negative away from zero
      -ceiling(-number, multiple, 0)
    end
  end
  
end
