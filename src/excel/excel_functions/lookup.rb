module ExcelFunctions
  
  # Looks like can be implemented as a MATCH(lookup_value, lookup_vector, 1) and then INDEX on either the lookup_vector or the result_vector
  # If used with lookup_vector in array form, first need to extract the first and last column, then do the above.
  def lookup(lookup_value, lookup_vector, result_vector = nil)
    return lookup_value if lookup_value.is_a?(Symbol)
    return lookup_vector if lookup_vector.is_a?(Symbol)
    return result_vector if result_vector.is_a?(Symbol)
    # Check if we are in array form
    if result_vector == nil && lookup_vector.length > 1 && lookup_vector.first.length > 1
      first_column = []
      last_column = []
      lookup_vector.each do |row|
        first_column.push(row.first)
        last_column.push(row.last)
      end
      return lookup(lookup_value, [first_column], [last_column])
    end
    i = excel_match(lookup_value, lookup_vector, 1)
    return i if i.is_a?(Symbol)
    return index(result_vector || lookup_vector, i)
  end
  
end
