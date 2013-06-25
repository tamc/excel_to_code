class ReplaceColumnWithColumnNumberAST
    
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
    if name == "COLUMN" && args.size == 1 && [:cell, :sheet_reference].include?(args[0][0])
      if args[0][0] == :cell
        reference = Reference.for(args[0][1])
      elsif args[0][0] == :sheet_reference
        reference = Reference.for(args[0][2][1])
      end
      reference.calculate_excel_variables
      @replacements_made_in_the_last_pass += 1
      [:number, reference.excel_column_number.to_s]
    else
      [:function,name,*args.map { |a| map(a) }]
    end
  end

end
  

class ReplaceColumnWithColumnNumber
    
  def self.replace(*args)
    self.new.replace(*args)
  end

  attr_accessor :replacements_made_in_the_last_pass
  
  def replace(input,output)
    rewriter = ReplaceColumnWithColumnNumberAST.new
    input.each_line do |line|
      # Looks to match lines with references
      if line =~ /"COLUMN"/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
    @replacements_made_in_the_last_pass = rewriter.replacements_made_in_the_last_pass
  end
end
