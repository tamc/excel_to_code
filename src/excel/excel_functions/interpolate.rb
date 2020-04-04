module ExcelFunctions
  
  def interpolate(r, i)
    return r if r.is_a?(Symbol)
    i = ensure_is_number(i)
    return i if i.is_a?(Symbol)

    r = [[r]] unless r.is_a?(Array)
    r.flatten!

    return :value if i < 1
    return :value if i > r.size

    # -1 because Excel arrays are 1 indexed, but ruby arrays are 0 indexed
    i = i - 1

    lower_i = i.floor
    higher_i = i.ceil
    i_fraction = i - lower_i

    return r[i] if i_fraction == 0

    lower_v = r[lower_i]
    higher_v = r[higher_i]
    v_range = higher_v - lower_v

    return lower_v + (i_fraction * v_range)
  end
  
end
