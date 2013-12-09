require_relative '../excel/formula_peg'

class ReplaceIndirectsWithReferencesAst
  
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
    return unless ast[1] == :INDIRECT
    args = ast[2..-1]
    if args[0][0] == :string
      @count_replaced += 1
      @replacement_made = true
      ast.replace(CachingFormulaParser.parse(args[0][1]))
    elsif args[0][0] == :error
      @count_replaced += 1
      @replacement_made = true
      ast.replace(args[0])
    end
  end
end
  

class ReplaceIndirectsWithReferences
    
  def self.replace(*args)
    self.new.replace(*args)
  end

  attr_accessor :count_replaced
  
  def replace(input,output)
    rewriter = ReplaceIndirectsWithReferencesAst.new
    input.each_line do |line|
      # Looks to match lines with references
      if line =~ /:INDIRECT/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
    @count_replaced = rewriter.count_replaced
  end
end
