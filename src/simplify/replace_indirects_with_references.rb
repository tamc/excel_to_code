require_relative '../excel/formula_peg'

class ReplaceIndirectsWithReferencesAst
  
  attr_accessor :replacements_made_in_the_last_pass

  def initialize
    @replacements_made_in_the_last_pass = 0
  end
   
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
      @replacements_made_in_the_last_pass += 1
      Formula.parse(args[0][1]).to_ast[1]
    elsif name == "INDIRECT" && args.size == 1 && args[0][0] == :error
      @replacements_made_in_the_last_pass += 1
      args[0]
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

  attr_accessor :replacements_made_in_the_last_pass
  
  def replace(input,output)
    rewriter = ReplaceIndirectsWithReferencesAst.new
    input.each_line do |line|
      # Looks to match lines with references
      if line =~ /"INDIRECT"/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
    @replacements_made_in_the_last_pass = rewriter.replacements_made_in_the_last_pass
  end
end
