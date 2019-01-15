class ReplaceCellAddressesWithReferencesAst

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
    _, name, info, ref = *ast
    return unless name == :CELL
    return unless info.first == :string
    return unless info.last.downcase == "address"
    replace_with_reference(ast, ref)
  end

  def replace_with_reference(ast, ref)
    case ref.first
    when :cell
      ast.replace([:cell, Reference.for(ref.last).unfix])
    when :sheet_reference
      replace_with_reference(ref.last, ref.last)
      ast.replace(ref)
    when :area
      ast.replace([:cell, Reference.for(ref[1]).unfix])
    else
      ast.replace([:error, "#VALUE!"])
    end
  end

end
  

class ReplaceCellAddressesWithReferences

    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  attr_accessor :count_replaced
  
  def replace(input,output)
    rewriter = ReplaceCellAddressesWithReferencesAst.new
    input.each_line do |line|
      # Looks to match lines with cell
      if line =~ /:CELL/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
    @count_replaced = rewriter.count_replaced
  end
end
