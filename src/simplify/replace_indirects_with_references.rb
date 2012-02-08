require_relative '../excel/formula_peg'

class ReplaceIndirectsWithReferencesAst
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if respond_to?(operator)
      send(operator,*ast[1..-1])
    else
      [operator,*ast[1..-1].map {|a| map(a) }]
    end
  end
  
  def function(name,*args)
    if name == "INDIRECT" && args.size == 1 && args[0][0] == :string
      Formula.parse(args[0][1]).to_ast[1]
    else
      puts "indirect #{[:function,name,*args.map { |a| map(a) }].inspect} not replaced" if name == "INDIRECT"
      [:function,name,*args.map { |a| map(a) }]
    end
  end
end
  

class ReplaceIndirectsWithReferences
    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    rewriter = ReplaceIndirectsWithReferencesAst.new
    input.lines do |line|
      # Looks to match lines with references
      if line =~ /"INDIRECT"/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
  end
end
