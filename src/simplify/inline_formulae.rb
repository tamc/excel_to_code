class Array

  def original
    @original || self
  end

  alias_method :original_replace, :replace

  def replace(new_array)
    @original = self.dup
    original_replace(new_array)
  end
end



class InlineFormulaeAst

  attr_accessor :references, :current_sheet_name, :inline_ast, :named_references
  attr_accessor :count_replaced
  
  def initialize(references = nil, current_sheet_name = nil, inline_ast = nil, named_references = {})
    @references, @current_sheet_name, @inline_ast, @named_references = references, [current_sheet_name], inline_ast, named_references
    @count_replaced = 0
    @inline_ast ||= lambda { |sheet, ref, references| true } # Default is to always inline
  end
  
  def map(ast)
    return ast unless ast.is_a?(Array)
    case ast[0]
    when :function
      function(ast)
    when :sheet_reference
      sheet_reference(ast)
    when :cell
      cell(ast)
    else
      ast.each do |a| 
        map(a) if a.is_a?(Array)
      end
    end
    ast
  end

  def function(ast)
    case ast[1]
    when :OFFSET
      # Don't map the second argument - it should be left as a cell refernce
      if (ast[2][0] == :cell || ast[2][0] == :sheet_reference)
        ast[3..-1].each {|a| map(a) }
      else
        ast.each { |a| map(a) }
      end
    when :COLUMN, :ROW
      # Don't map any arguments
    else
      # Otherwise good to map all the other arguments
      ast.each { |a| map(a) }
    end
  end
  
  # Should be of the form [:sheet_reference, sheet_name, reference]
  # FIXME: Can we rely on reference always being a [:cell, ref] at this stage?
  # FIXME: NO! Because they won't be when they are used in EmergencyArrayFormulaReplaceIndirectBodge
  def sheet_reference(ast)
    return unless ast[2][0] == :cell
    sheet = ast[1].to_sym
    ref = ast[2][1].to_s.upcase.gsub('$','').to_sym
    ref_org = ast[2][1].to_s.gsub('$','').to_sym

    if named_references[ref_org] # Fix issue for named_reference that is mapped to blank
      sheet = named_references[ref_org][1]
      ref = named_references[ref_org][2][1]
    else
      # FIXME: Need to check if valid worksheet and return [:error, "#REF!"] if not
      # Now check user preference on this
      return unless inline_ast.call(sheet,ref, references)
    end

    ast_to_inline = ast_or_blank(sheet, ref)
    @count_replaced += 1
    current_sheet_name.push(sheet)
    map(ast_to_inline)
    current_sheet_name.pop
    ast.replace(ast_to_inline)
  end
  
  # Format [:cell, ref]
  def cell(ast)
    sheet = current_sheet_name.last
    ref = ast[1].to_s.upcase.gsub('$', '').to_sym
    if inline_ast.call(sheet, ref, references)
      ast_to_inline = ast_or_blank(sheet, ref)
      @count_replaced += 1
      map(ast_to_inline)
      ast.replace(ast_to_inline)
    # FIXME: Check - is this right? does it work recursively enough?
    elsif current_sheet_name.size > 1 
      ast.replace([:sheet_reference, sheet, ast.dup])
    end
  end

  def ast_or_blank(sheet, ref)
    ast_to_inline = references[[sheet, ref]]
    return ast_to_inline if ast_to_inline

    if named_references.key?(ref.downcase)
      return named_references[ref.downcase]
    end

    # Need to add a new blank cell and return ast for an inlined blank
    references[[sheet, ref]] = [:blank]
    [:inlined_blank]
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
