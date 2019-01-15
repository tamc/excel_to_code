module ExcelFunctions
  
  # This is a support function for mapping arguments that are not strings to their
  # Excel string equivalent. e.g., true = TRUE, :div0 = "DIV/0!"
  def string_argument(a)
    case a
    when Symbol
      return {
        :name => "#NAME?",
        :value => "#VALUE!",
        :div0 => "#DIV/0!",
        :ref => "#REF!",
        :na => "#N/A",
        :num => "#NUM!",
      }[a] || :value
    when String
      return a
    when nil
      return "" 
    when true
      return "TRUE"
    when false
      return "FALSE" 
    when Numeric
      if a.round == a
        return a.to_i.to_s
      else
        return a.to_s
      end
    when Array
      return string_argument(a[0][0])
    else
      return :value
    end
  end
  
end
