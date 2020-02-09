require_relative '../excel/formula_peg'

class EmergencyArrayFormulaReplaceIndirectBodge
  
  attr_accessor :current_sheet_name
  attr_accessor :references
  attr_accessor :named_references
  attr_accessor :tables
  attr_accessor :referring_cell

  def initialize
    @import_replacer = ReplaceImportWithReference.new
    @indirect_replacer = ReplaceIndirectsWithReferencesAst.new
    @formulae_to_value_replacer = MapFormulaeToValues.new
    @inline_formulae_replacer = InlineFormulaeAst.new
    @simplify_arithmetic_replacer ||= SimplifyArithmeticAst.new
    @replace_ranges_with_array_literals_replacer ||= ReplaceRangesWithArrayLiteralsAst.new
  end

  def replace(ast)
    map(ast)
    ast
  end
   
  def map(ast)
    return ast unless ast.is_a?(Array)
    function(ast) if ast[0] == :function && (ast[1] == :INDIRECT || ast[1] == :IMPORT)
    ast.each { |a| map(a) }
    ast
  end
  
  def function(ast)
    new_ast = deep_copy(ast)
    @replace_ranges_with_array_literals_replacer.map(new_ast)
    @inline_formulae_replacer.references = @references
    @inline_formulae_replacer.current_sheet_name = [@current_sheet_name]
    @inline_formulae_replacer.map(new_ast)
    @simplify_arithmetic_replacer.map(new_ast)

    @formulae_to_value_replacer.map(new_ast)
    
    @named_reference_replacer ||= ReplaceNamedReferencesAst.new(@named_references)
    @named_reference_replacer.default_sheet_name = @current_sheet_name
    @named_reference_replacer.map(new_ast)
    
    @table_reference_replacer ||= ReplaceTableReferenceAst.new(@tables)
    @table_reference_replacer.worksheet = @current_sheet_name
    @table_reference_replacer.referring_cell = @referring_cell
    @table_reference_replacer.map(new_ast)

    @replace_ranges_with_array_literals_replacer.map(new_ast)
    @inline_formulae_replacer.map(new_ast)
    @formulae_to_value_replacer.map(new_ast)
    
    @import_replacer.replace(new_ast)
    @indirect_replacer.replace(new_ast)
    
    ast.replace(new_ast)
    ast
  end

  def deep_copy(ast)
    return ast if ast.is_a?(Symbol)
    return ast if ast.is_a?(Numeric)
    return ast.dup unless ast.is_a?(Array)
    ast.map do |a|
      deep_copy(a)
    end
  end
end
