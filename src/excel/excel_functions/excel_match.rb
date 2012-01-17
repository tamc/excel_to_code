module ExcelFunctions
  
  def excel_match(lookup_value,lookup_array,match_type = 0)
    return lookup_value if lookup_value.is_a?(Symbol)
    return lookup_array if lookup_array.is_a?(Symbol)
    return match_type if match_type.is_a?(Symbol)
    
    lookup_value ||= 0
    lookup_array ||= [0]
    match_type ||= 0
    lookup_array = [lookup_array] unless lookup_array.is_a?(Array)
    lookup_array.flatten!
    
    lookup_value = lookup_value.downcase if lookup_value.respond_to?(:downcase)
    case match_type      
    when 0, 0.0, false
      lookup_array.each_with_index do |item,index|
        item ||= 0
        item = item.downcase if item.respond_to?(:downcase)
        return index+1 if lookup_value == item
      end
      return :na
    when 1, 1.0, true
      lookup_array.each_with_index do |item, index|
        item ||= 0
        next if lookup_value.is_a?(String) && !item.is_a?(String)
        next if lookup_value.is_a?(Numeric) && !item.is_a?(Numeric)
        item = item.downcase if item.respond_to?(:downcase)
        if item > lookup_value
          return :na if index == 0
          return index
        end
      end
      return lookup_array.to_a.size
    when -1, -1.0
      lookup_array.each_with_index do |item, index|
        item ||= 0
        next if lookup_value.is_a?(String) && !item.is_a?(String)
        next if lookup_value.is_a?(Numeric) && !item.is_a?(Numeric)
        item = item.downcase if item.respond_to?(:downcase)
        if item < lookup_value
          return :na if index == 0
          return index
        end
      end
      return lookup_array.to_a.size - 1
    end
    return :na
  end
  
end
