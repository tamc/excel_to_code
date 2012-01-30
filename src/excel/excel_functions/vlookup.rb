module ExcelFunctions
  
  def vlookup(lookup_value,lookup_table,column_number, match_type = true)    
    return lookup_value if lookup_value.is_a?(Symbol)
    return lookup_table if lookup_table.is_a?(Symbol)
    return column_number if column_number.is_a?(Symbol)
    return match_type if match_type.is_a?(Symbol)
    
    return :na if lookup_value == nil
    return :na if lookup_table == nil
    return :na if column_number == nil
    return :na if match_type == nil
  
    lookup_value = lookup_value.downcase if lookup_value.is_a?(String)
    
    last_good_match = 0
    
    lookup_table.each_with_index do |row,index|
      possible_match = row.first
      
      next if lookup_value.is_a?(String) && !possible_match.is_a?(String)
      next if lookup_value.is_a?(Numeric) && !possible_match.is_a?(Numeric)
      
      possible_match.downcase! if lookup_value.is_a?(String)

      if lookup_value == possible_match
        return :value unless column_number <= row.length
        return row[column_number-1]
      elsif match_type == true
        if possible_match > lookup_value
          return :na if index == 0
          previous_row = lookup_table[last_good_match]
          return :value unless column_number <= previous_row.length
          return previous_row[column_number-1]
        else
          last_good_match = index
        end
      end      
    end

    # We don't have a match
    if match_type == true
      return lookup_table[last_good_match][column_number-1]
    else
      return :na
    end
  end
  
end
