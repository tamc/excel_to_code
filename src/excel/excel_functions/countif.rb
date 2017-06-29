module ExcelFunctions
  
  def countif(count_range, *criteria)
    # First, get rid of the errors
    return count_range if count_range.is_a?(Symbol)
    error = criteria.find { |a| a.is_a?(Symbol) }
    return error if error
    
    # Sort out the count range
    count_range = [count_range] unless count_range.is_a?(Array)
    count_range = count_range.flatten
    
    # Sort out the criteria
    0.step(criteria.length-1,2).each do |i|
      if criteria[i].is_a?(Array)
        criteria[i] = criteria[i].flatten
      else
        criteria[i] = [criteria[i]]
      end
    end
      
    # This will be the final answer
    accumulator = 0
    
    # Work through each part of the count range
    count_range.each_with_index do |potential_count,index|
      next unless potential_count.is_a?(Numeric)
      
      # If a criteria fails, this is set to false and no further criteria are evaluated
      pass = true
      
      0.step(criteria.length-1,2).each do |i|
        check_range = criteria[i]
        required_value = criteria[i+1]
        return :value if index >= check_range.length
        check_value = check_range[index]
        
        pass = case check_value
        when String
          check_value.downcase == required_value.to_s.downcase
        when true, false
          check_value == required_value
        when nil
          required_value == ""
        when Numeric
          case required_value
          when Numeric
            check_value == required_value
          when String
            required_value =~ /^(<=|>=|<|>)?([-+]?[0-9]+\.?[0-9]*([eE][-+]?[0-9]+)?)$/
            if $1 && $2
              check_value.send($1,$2.to_f)
            elsif $2
              check_value == $2.to_f
            else
              false
            end
          else
            check_value == required_value
          end
        end # case check_value
                
        break unless pass
      end # criteria loop
      
      accumulator += 1 if pass
        
      end
      
      return accumulator
      
    end
  
end
