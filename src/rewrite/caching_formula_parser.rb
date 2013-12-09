require 'singleton'

class CachingFormulaParser
  include Singleton

  def self.parse(*args)
    instance.parse(*args)
  end

  def initialize
    @number_cache = {}
    @string_cache = {}
  end

  def parse(text)
    ast = Formula.parse(text)
    if ast
      map(ast.to_ast[1])
    else
      nil
    end
  end

  def map(ast)
    return ast unless ast.is_a?(Array)
    if respond_to?(ast[0])
      ast = send(ast[0], ast) 
    else
      ast.map! { |a| map(a) }
    end
    ast
  end

  def sheet_reference(ast)
    ast[1] = ast[1].to_sym
    ast[2] = map(ast[2])
    ast
  end

  def cell(ast)
    ast[1] = ast[1].to_sym
    ast
  end

  def area(ast)
    ast[1] = ast[1].to_sym
    ast[2] = ast[2].to_sym
    ast
  end

  def number(ast)
    ast[1] = ast[1].to_f
    @number_cache[ast] ||= ast
  end

  def string(ast)
    return @string_cache[ast] ||= ast
  end

end
