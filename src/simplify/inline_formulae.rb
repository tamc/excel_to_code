class InlineFormulaeAst
  
  attr_accessor :references, :current_sheet_name
  
  def initialize(references, current_sheet_name)
    @references, @current_sheet_name = references, [current_sheet_name]
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
  
  # TODO: Optimize by replacing contents of references hash with the inlined version
  def sheet_reference(sheet,reference)
    ast = references[sheet][reference.last.gsub('$','')]
    if ast
      current_sheet_name.push(sheet)
      result = map(ast)
      current_sheet_name.pop
    else
      result = [:sheet_reference,sheet,reference]
    end
    result  
  end
  
  # TODO: Optimize by replacing contents of references hash with the inlined version
  def cell(reference)
    ast = references[current_sheet_name.last][reference.gsub('$','')]
    if ast
      map(ast)
    else
      [:cell,reference]
    end
  end
    
end
  

class InlineFormulae
  
  attr_accessor :references, :default_sheet_name
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    rewriter = InlineFormulaeAst.new(references,default_sheet_name)
    input.lines do |line|
      # Looks to match lines with references
      if line =~ /\[:cell/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
  end
end
