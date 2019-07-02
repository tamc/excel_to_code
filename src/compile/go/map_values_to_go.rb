# frozen_string_literal: true

require_relative '../../util/not_supported_exception'

class MapValuesToGo

  def map(ast)
    raise NotSupportedException, "#{ast} not supported" unless ast.is_a?(Array)

    operator = ast[0]
    raise NotSupportedException, "#{operator} in #{ast} not supported" unless respond_to?(operator)

    send(operator, *ast[1..-1])
  end

  def blank
    @result&.blank!
    'Blank{}'
  end

  def inlined_blank
    @result&.blank!
    'Blank{}'
  end

  def constant(name)
    name
  end

  alias null blank

  def number(text)
    @result&.number!
    text.to_s
  end

  def percentage(text)
    @result&.number!
    (text.to_f / 100.0).to_s
  end

  def string(text)
    @result&.string!
    text.inspect
  end

  ERRORS = {
    "#NAME?": 'NameError{}',
    "#VALUE!": 'ValueError{}',
    "#DIV/0!": 'Div0Error{}',
    "#REF!": 'RefError{}',
    "#N/A": 'NAError{}',
    "#NUM!": 'NumError{}'
  }.freeze

  REVERSE_ERRORS = ERRORS.invert

  def error(text)
    @result&.error!
    ERRORS[text.to_sym] || (raise NotSupportedException, "#{text.inspect} error not recognised")
  end

  def boolean_true
    @result&.boolean!
    'true'
  end

  def boolean_false
    @result&.boolean!
    'false'
  end
end
