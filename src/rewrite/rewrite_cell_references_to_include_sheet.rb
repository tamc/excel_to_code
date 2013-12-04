require_relative '../excel'

class RewriteCellReferencesToIncludeSheetAst 
  
  attr_accessor :worksheet
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    if respond_to?(ast[0])
      send(ast[0], ast)    
    else 
      ast.each { |a| map(a) }
    end
    ast
  end
  
  def cell(ast)
    ast.replace([:sheet_reference, worksheet, ast.dup])
  end
  
  def area(ast)
    ast.replace([:sheet_reference, worksheet, ast.dup])
  end
  
  def sheet_reference(ast)
    # Leave alone, don't map futher
  end

end

class RewriteCellReferencesToIncludeSheet
  
  def self.rewrite(*args)
    new.rewrite(*args)
  end
  
  attr_accessor :worksheet
  
  def rewrite(input,output)
    mapper = RewriteCellReferencesToIncludeSheetAst.new
    mapper.worksheet = worksheet
    input.each_line do |line|
      if line =~ /(:area|:cell)/
        content = line.split("\t")
        ast = eval(content.pop)
        output.puts "#{content.join("\t")}\t#{mapper.map(ast).inspect}"
      else
        output.puts line
      end
    end
  end
  
end
