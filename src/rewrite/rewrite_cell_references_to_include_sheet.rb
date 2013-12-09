require_relative '../excel'

class RewriteCellReferencesToIncludeSheetAst 
  
  attr_accessor :worksheet

  def initialize
    @fp = CachingFormulaParser.instance
  end
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    return cell(ast) if ast[0] == :cell
    return area(ast) if ast[0] == :area
    return sheet_reference(ast) if ast[0] == :sheet_reference
    ast.each.with_index do |a,i|
      next unless a.is_a?(Array)
      case a[0]
      when :cell
        ast[i] = cell(a)
      when :area
        ast[i] = area(a)
      when :sheet_reference
        ast[i] = sheet_reference(a)
      else
        map(a)
      end
    end
    ast
  end
  
  def cell(ast)
    @fp.map([:sheet_reference, worksheet, ast.dup])
  end
  
  def area(ast)
    @fp.map([:sheet_reference, worksheet, ast.dup])
  end
  
  def sheet_reference(ast)
    @fp.map(ast)
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
