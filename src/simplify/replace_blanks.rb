class ReplaceBlanksAst
  
  attr_accessor :references, :default_sheet_name
  
  def initialize(references, default_sheet_name)
    @references, @default_sheet_name = references, default_sheet_name
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
  
  def sheet_reference(sheet,reference)
    if references[sheet].has_key?(reference.last.gsub('$',''))
      [:sheet_reference,sheet,reference]
    else
      [:blank]
    end
  end
  
  def cell(reference)
    if references[default_sheet_name].has_key?(reference.gsub('$',''))
      [:cell,reference]
    else
      [:blank]
    end
  end
  
end
  

class ReplaceBlanks
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,references,sheet_name,output)
    rewriter = ReplaceBlanksAst.new(references,sheet_name)
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
