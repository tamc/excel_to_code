require_relative '../../util/not_supported_exception'

class MapValuesToGo

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
  
  def blank
    "excel.Blank{}"
  end

  def inlined_blank
    "excel.Blank{}"
  end
  
  def constant(name)
    name
  end
  
  alias :null :blank
    
  def number(text)
    text.to_s
  end
  
  def percentage(text)
    (text.to_f / 100.0).to_s
  end
  
  def string(text)
    text.inspect
  end
  
  ERRORS = {
    :"#NAME?" => "excel.NameError{}",
    :"#VALUE!" => "excel.ValueError{}",
    :"#DIV/0!" => "excel.Div0Error{}",
    :"#REF!" => "excel.RefError{}",
    :"#N/A" => "excel.NAError{}",
    :"#NUM!" => "excel.NumError{}"
  }
  
  REVERSE_ERRORS = ERRORS.invert
  
  def error(text)
    ERRORS[text.to_sym] || (raise NotSupportedException.new("#{text.inspect} error not recognised"))
  end
  
  def boolean_true
    "true"
  end
  
  def boolean_false
    "false"
  end
  
end
