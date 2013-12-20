module ExcelFunctions
  
  def hlookup(lookup_value, lookup_table, row_number, match_type = true)    
    return lookup_value if lookup_value.is_a?(Symbol)
    return lookup_table if lookup_table.is_a?(Symbol)
    return row_number if row_number.is_a?(Symbol)
    return match_type if match_type.is_a?(Symbol)
    
    return :na if lookup_value == nil
    return :na if lookup_table == nil
    return :na if row_number == nil
    return :na if match_type == nil
  
    lookup_value = lookup_value.downcase if lookup_value.is_a?(String)
    
    last_good_match = 0

    return :value unless row_number > 0
    return :ref unless row_number <= lookup_table.size

    lookup_table.first.each_with_index do |possible_match, column_number|
      
      next if lookup_value.is_a?(String) && !possible_match.is_a?(String)
      next if lookup_value.is_a?(Numeric) && !possible_match.is_a?(Numeric)
      
      possible_match = possible_match.downcase if lookup_value.is_a?(String)

      if lookup_value == possible_match
        return lookup_table[row_number-1][column_number]
      elsif match_type == true
        if possible_match > lookup_value
          return :na if column_number == 0
          return lookup_table[row_number-1][last_good_match]
        else
          last_good_match = column_number
        end
      end      
    end

    # We don't have a match
    if match_type == true
      return lookup_table[row_number - 1][last_good_match]
    else
      return :na
    end
  end
  
end
