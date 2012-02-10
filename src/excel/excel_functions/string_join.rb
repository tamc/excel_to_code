module ExcelFunctions
  
  def string_join(*strings)
    strings.find {|s| s.is_a?(Symbol)} || strings.map do |s| 
      case s
      when nil; "0"
      when Numeric
        if s.round == s
          s.to_i.to_s
        else
          s.to_s
        end
      else
        s.to_s
      end
    end.join('')
  end
  
end
