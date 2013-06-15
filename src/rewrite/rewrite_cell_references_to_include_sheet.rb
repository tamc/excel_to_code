require_relative '../excel'

class RewriteCellReferencesToIncludeSheetAst 
  
  attr_accessor :worksheet
    
  def map(ast)
    if ast.is_a?(Array)
      operator = ast.shift
      if respond_to?(operator)
        send(operator,*ast)
      else
        [operator,*ast.map {|a| map(a) }]
      end
    else
      return ast
    end
  end
  
  def cell(ref)
    [:sheet_reference, worksheet, [:cell, ref]]
  end
  
  def area(start,finish)
    [:sheet_reference, worksheet, [:area, start, finish]]    
  end
  
  def sheet_reference(sheet_name,reference)
    [:sheet_reference, sheet_name, reference]
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
