require_relative '../excel/formula_peg'

class EmergencyArrayFormulaReplaceIndirectBodge
  
  attr_accessor :current_sheet_name
  attr_accessor :references

  def initialize
    @indirect_replacer = ReplaceIndirectsWithReferencesAst.new
    @formulae_to_value_replacer = MapFormulaeToValues.new
    @inline_formulae_replacer = InlineFormulaeAst.new
  end

  def replace(ast)
    map(ast)
    ast
  end
   
  def map(ast)
    return ast unless ast.is_a?(Array)
    function(ast) if ast[0] == :function && ast[1] == :INDIRECT
    ast.each { |a| map(a) }
    ast
  end
  
  def function(ast)
    new_ast = deep_copy(ast)
    @inline_formulae_replacer.references = @references
    @inline_formulae_replacer.current_sheet_name = [@current_sheet_name]
    @inline_formulae_replacer.map(new_ast)
    @formulae_to_value_replacer.map(new_ast)
    @indirect_replacer.replace(new_ast)
    ast.replace(new_ast)
    ast
  end

  def deep_copy(ast)
    return ast if ast.is_a?(Symbol)
    return ast.dup unless ast.is_a?(Array)
    ast.map do |a|
      deep_copy(a)
    end
  end
end
