require_relative '../../util/not_supported_exception'

class MapValuesToRuby

  def map(ast)
    if ast.is_a?(Array)
      operator = ast.shift
      if respond_to?(operator)
        send(operator,*ast)
      else
        raise NotSupportedException.new("#{operator} in #{ast.inspect} not supported")
      end
    else
      raise NotSupportedException.new("#{ast} not supported")
    end
  end
    
  def number(text)
    case text
    when /\./
      text.to_f
    when /e/i
      text.to_f
    else
      text.to_i
    end
  end
  
  def string(text)
    text
  end
  
  ERRORS = {
    "#NAME?" => :name
  }
  
  def error(text)
    ERRORS[text] || (raise NotSupportedException.new("#{text} error not recognised"))
  end
  
  def boolean_true
    true
  end
  
  def boolean_false
    false
  end
  
end
