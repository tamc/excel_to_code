require_relative '../excel'

class RewriteCellReferencesToIncludeSheetAst 
  
  attr_accessor :worksheet

  def initialize
    @fp = CachingFormulaParser.instance
  end
    
  # FIXME: UGh.
  def map(ast)
    r = do_map(ast)
    ast.replace(r) unless r.object_id == ast.object_id
    ast
  end

  def do_map(ast)
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
        do_map(a)
      end
    end
    ast
  end
  
  def cell(ast)
    ast[1] = ast[1].to_s.gsub('$','').to_sym
    @fp.map([:sheet_reference, worksheet, ast.dup])
  end
  
  def area(ast)
    ast[1] = ast[1].to_s.gsub('$','').to_sym
    ast[2] = ast[2].to_s.gsub('$','').to_sym
    @fp.map([:sheet_reference, worksheet, ast.dup])
  end
  
  def sheet_reference(ast)
    if ast[2][0] == :cell
      ast[2][1] = ast[2][1].to_s.gsub('$','').to_sym
    elsif ast[2][1] == :area
      ast[2][1] = ast[2][1].to_s.gsub('$','').to_sym
      ast[2][2] = ast[2][2].to_s.gsub('$','').to_sym
    end
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
