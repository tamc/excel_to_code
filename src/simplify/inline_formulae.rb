class InlineFormulaeAst
  
  attr_accessor :references, :current_sheet_name, :inline_ast
  attr_accessor :count_replaced
  
  def initialize(references = nil, current_sheet_name = nil, inline_ast = nil)
    @references, @current_sheet_name, @inline_ast = references, [current_sheet_name], inline_ast
    @count_replaced = 0
    @inline_ast ||= lambda { |sheet, ref, references| true } # Default is to always inline
  end
  
  def map(ast)
    return ast unless ast.is_a?(Array)
    if respond_to?(ast[0])
      send(ast[0], ast) 
    else # In this case needs to be an else because don't want to map first argument in OFFSET(cell_to_offset_from_shouldn't_be_mapped, rows, columns)
      ast.each { |a| map(a) }
    end
    ast
  end

  def function(ast)
    if ast[1] == 'OFFSET'
      # Don't map the second argument - it should be left as a cell refernce
      ast[3..-1].each {|a| map(a) }
    else
      # Otherwise good to map all the other arguments
      ast.each { |a| map(a) }
    end
  end
  
  # Should be of the form [:sheet_reference, sheet_name, reference]
  # FIXME: Can we rely on reference always being a [:cell, ref] at this stage?
  def sheet_reference(ast)
    return unless ast[2][0] == :cell
    sheet = ast[1]
    ref = ast[2][1].upcase.gsub('$','')
    # FIXME: Need to check if valid worksheet and return [:error, "#REF!"] if not
    # Now check user preference on this
    return unless inline_ast.call(sheet,ref, references)
    @count_replaced += 1
    ast_to_inline = references[[sheet, ref]]
    return ast.replace([:blank]) unless ast_to_inline
    current_sheet_name.push(sheet)
    map(ast_to_inline)
    current_sheet_name.pop
    ast.replace(ast_to_inline)
  end
  
  # Format [:cell, ref]
  def cell(ast)
    sheet = current_sheet_name.last
    ref = ast[1].upcase.gsub('$', '')
    if inline_ast.call(sheet, ref, references)
      @count_replaced += 1
      ast_to_inline = references[[sheet, ref]]
      return ast.replace([:blank]) unless ast_to_inline
      map(ast_to_inline)
      ast.replace(ast_to_inline)
    # FIXME: Check - is this right? does it work recursively enough?
    elsif current_sheet_name.size > 1 
      ast.replace([:sheet_reference, sheet, ast.dup])
    end
  end
    
end
  

class InlineFormulae
  
  attr_accessor :references, :default_sheet_name, :inline_ast
  
  def self.replace(*args)
    self.new.replace(*args)
  end

  attr_accessor :count_replaced
  
  def replace(input,output)
    rewriter = InlineFormulaeAst.new(references, default_sheet_name, inline_ast)
    input.each_line do |line|
      # Looks to match lines with references
      if line =~ /\[:cell/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
    @count_replaced = rewriter.count_replaced
  end
end
