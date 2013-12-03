class CountFormulaReferences
  
  # FIXME: Do we need these accessors
  attr_accessor :references
  attr_accessor :dependencies
  attr_accessor :current_sheet
  
  def initialize(references = {}, dependencies = {})
    @references = references
    @dependencies = dependencies
    dependencies.default_proc = lambda do |hash,key|
      hash[key] = 0
    end
    @current_sheet = []
  end
  
  def count(references)
    # FIXME: Why do we have this references instance variable?
    @references = references
    references.each do |full_ref,ast|
      @dependencies[full_ref] ||= 0
      count_dependencies_for(full_ref.first,full_ref.last,ast)
    end
    return dependencies
  end
    
  def count_dependencies_for(sheet,ref,ast)
    current_sheet.push(sheet)
    map(ast)
    current_sheet.pop
  end
  
  def map(ast)
    return ast unless ast.is_a?(Array)
    case ast[0]
    when :sheet_reference; sheet_reference(ast)
    when :cell; cell(ast)
    else; ast.each { |a| map(a) }
    end
    ast
  end
  
  # Format [:sheet_reference, sheet, reference]
  def sheet_reference(ast)
    sheet = ast[1]
    reference = ast[2]
    ref = reference.last.gsub('$','')
    @dependencies[[sheet, ref]] += 1
  end
  
  # Format [:cell, reference]
  def cell(ast)
    reference = ast[1]
    ref = reference.gsub('$','')
    @dependencies[[current_sheet.last, ref]] += 1
  end
   
end
