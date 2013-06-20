module ExcelFunctions
  
  def pv(rate, nper, pmt, fv = nil, type = nil)
    # Turn the remainder into numbers
    rate = number_argument(rate)
    nper = number_argument(nper)
    pmt = number_argument(pmt)
    fv = number_argument(fv)
    type = number_argument(type)

    # Check for errors
    return rate if rate.is_a?(Symbol)
    return nper if nper.is_a?(Symbol)
    return pmt if pmt.is_a?(Symbol)
    return fv if fv.is_a?(Symbol)

    return :value unless (type == 1 || type == 0)
    return :value unless rate >= 0

    # Sum the payments
    if rate > 0
      present_value = -pmt * ((1 - ((1 + rate)**-nper))/rate)
    else
      present_value = -pmt * nper
    end

    # Adjust for the type, which governs whether payments at the beginning or end of the period
    present_value = present_value * ( 1 + rate) if type == 1

    # Add on the final value
    present_value += -fv / ((1 + rate)**(nper))

    # Return the answer
    present_value
  end
  
end
