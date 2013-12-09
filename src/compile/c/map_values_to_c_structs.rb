require_relative '../../util/not_supported_exception'

class MapValuesToCStructs

  def map(ast)
    if ast.is_a?(Array)
      operator = ast[0]
      if respond_to?(operator)
        send(operator,*ast[1..-1])
      else
        raise NotSupportedException.new("#{operator} in #{ast.inspect} not supported")
      end
    else
      raise NotSupportedException.new("#{ast} not supported")
    end
  end
  
  def number(text)
    n = case text.to_s
    when /\./
      text.to_f.to_s
    when /e/i
      text.to_f.to_s
    else
      text.to_i.to_s
    end
    "{.type = ExcelNumber, .number = #{n}}"
  end
  
  def percentage(text)
    "{.type = ExcelNumber, .number = #{(text.to_f / 100.0).to_s}}"
  end
  
  def string(text)
    "{.type = ExcelString, .string = #{text.inspect}}"
  end  
end
