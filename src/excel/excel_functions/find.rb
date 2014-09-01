module ExcelFunctions
  
  def find(find_text,within_text,start_number = 1)
    return find_text if find_text.is_a?(Symbol)
    return within_text if within_text.is_a?(Symbol)
    return start_number if start_number.is_a?(Symbol)

    # nils are treated as empty strings
    find_text ||= ""
    within_text ||= ""

    # Other things should be turned into a string
    find_text = find_text.to_s
    within_text = within_text.to_s
    
    # there are some cases where the start_number is remapped
    case start_number
    when nil; return :value
    when String; 
      begin
        start_number = Float(start_number)
      rescue ArgumentError
        return :value
      end
    when true; start_number = 1
    when false; return :value
    when Numeric; # ok
    else; return :value
    end
    
    # edge case
    return 1 if find_text == "" && within_text == "" && start_number == 1
    
    # length check
    return :value if start_number < 1
    return :value if start_number > within_text.length
      
    # Ok, lets go
    result = within_text.index(find_text,start_number - 1 )
    result ? result + 1 : :value
  end
  
end
