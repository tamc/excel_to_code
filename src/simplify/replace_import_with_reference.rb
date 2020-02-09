require_relative '../excel/formula_peg'

class ReplaceImportWithReference
  
  attr_accessor :count_replaced
  attr_accessor :replacement_made

  def initialize
    @count_replaced = 0
  end

  def replace(ast)
    @replacement_made = false
    map(ast)
    @replacement_made
  end
   
  def map(ast)
    return ast unless ast.is_a?(Array)
    function(ast) if ast[0] == :function
    ast.each { |a| map(a) }
    ast
  end
  
  def function(ast)
    return unless ast[1] == :IMPORT
    args = ast[2..-1]
    return unless args[0][0] == :string
    @count_replaced += 1
    @replacement_made = true
    name = args[0][1]
    prefix = "in"
    if args.length == 2 && args[1][0] == :string 
      prefix = args[1][1]
    end
    ast.replace([:named_reference, "#{prefix}.#{name}"])
  end
end
