require 'singleton'

class ExternalReferenceException < ExcelToCodeException

  attr_accessor :reference_ast
  attr_accessor :full_ast
  attr_accessor :formula_text
  attr_accessor :ref

  def initialize(reference_ast, full_ast, formula_text)
    @reference_ast, @full_ast, @formula_text = reference_ast, full_ast, formula_text
  end

  def message
    <<-END


    Sorry, ExcelToCode can't handle external references

    It found one in #{ref.join("!")}
    The formula was #{formula_text}
    Which was parsed to #{full_ast}
    Which seemed to have an external reference at #{reference_ast}
    Note, the [0], [1], [2] ...  are the way Excel stores the names of the external files. 
    
    Please remove the external reference from the Excel and try again.

    END
  end
end

class ParseFailedException < ExcelToCodeException

  attr_accessor :formula_text
  attr_accessor :ref

  def initialize(formula_text)
    @formula_text = formula_text
  end

  def message
    <<-END


    Sorry, ExcelToCode couldn't parse one of the formulae

    It was in #{ref.join("!")}
    The formula was #{formula_text}

    Please report the problem at http://github.com/tamc/excel_to_code/issues

    END
  end

end


class CachingFormulaParser
  include Singleton

  attr_accessor :functions_used

  def self.parse(*args)
    instance.parse(*args)
  end

  def self.map(*args)
    instance.map(*args)
  end

  def initialize
    @number_cache = {}
    @string_cache = {}
    @percentage_cache = {}
    @error_cache = {}
    @operator_cache = {}
    @comparator_cache = {}
    @sheet_reference_cache = {}
    @functions_used = {}
  end

  def parse(text)
    ast = Formula.parse(text)
    @text = text # Kept in case of Exception below
    if ast
      @full_ast = ast.to_ast[1] # Kept in case of Exception below
      map(ast.to_ast[1])
    else
      raise ParseFailedException.new(text)
      nil
    end
  end

  # FIXME: THe function bit in here isn't DRY or consistent
  def map(ast)
    return ast unless ast.is_a?(Array)
    if ast[0] == :function
      ast[1] = ast[1].to_sym 
      @functions_used[ast[1]] = true
    end
    if respond_to?(ast[0])
      ast = send(ast[0], ast) 
    else
      ast.each.with_index do |a,i| 
        next unless a.is_a?(Array)
        
        if a[0] == :function
          a[1] = a[1].to_sym 
          @functions_used[a[1]] = true
        end

        ast[i] = map(a)
      end
    end
    ast
  end

  # We can't deal with external references at the moment
  def external_reference(ast)
    raise ExternalReferenceException.new(ast, @full_ast, @text)
  end

  def sheet_reference(ast)
    # Sheet names shouldn't start with [1], because those are 
    # external references
    if ast[1] =~ /^\[\d+\]/
       raise ExternalReferenceException.new(ast, @full_ast, @text)
    end
    ast[1] = ast[1].to_sym
    ast[2] = map(ast[2])
    # We do this to deal with Control!#REF! style rerferences
    # that occasionally pop up in named references
    if ast[2].first == :error
      return ast[2]
    else
      @sheet_reference_cache[ast] ||= ast
    end
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

  def percentage(ast)
    if ast[1].is_a?(Array)
      ast.replace([:arithmetic, map(ast[1]), operator([:operator, :'/']), number([:number, 100])])
    else
      ast[1] = ast[1].to_f / 100.0
      ast[0] = :number
    end
    @percentage_cache[ast] ||= ast
  end

  def string(ast)
    return @string_cache[ast] ||= ast
  end

  TRUE = [:boolean_true]
  FALSE = [:boolean_false]
  BLANK = [:blank]

  def boolean_true(ast)
    TRUE
  end

  def boolean_false(ast)
    FALSE
  end

  def error(ast)
    ast[1] = ast[1].to_sym
    @error_cache[ast] ||= ast
  end

  def blank(ast)
    BLANK
  end

  def operator(ast)
    ast[1] = ast[1].to_sym
    @operator_cache[ast] ||= ast
  end

  def comparator(ast)
    ast[1] = ast[1].to_sym
    @comparator_cache[ast] ||= ast
  end

end
