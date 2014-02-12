module ExcelFunctions
  
  def sumifs(sum_range,*criteria)
    # First, get rid of the errors
    return sum_range if sum_range.is_a?(Symbol)
    error = criteria.find { |a| a.is_a?(Symbol) }
    return error if error
    
    # Sort out the sum range
    sum_range = [sum_range] unless sum_range.is_a?(Array)
    sum_range = sum_range.flatten
    
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
    
    # Work through each part of the sum range
    sum_range.each_with_index do |potential_sum,index|
      next unless potential_sum.is_a?(Numeric)
      
      # If a criteria fails, this is set to false and no further criteria are evaluated
      pass = true
      
      0.step(criteria.length-1,2).each do |i|
        check_range = criteria[i]
        required_value = criteria[i+1] || 0
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
      
      accumulator += potential_sum if pass
        
      end
      
      return accumulator
      
    end
  
end
