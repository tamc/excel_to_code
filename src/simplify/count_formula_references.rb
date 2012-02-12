class CountFormulaReferences
  
  attr_accessor :references
  attr_accessor :dependencies
  attr_accessor :current_sheet
  
  def initialize(references = {}, dependencies = {})
    @references = references
    @dependencies = dependencies
    dependencies.default_proc = lambda do |hash,key|
      hash[key] = {}
    end
    @current_sheet = []
  end
  
  def count(references)
    @references = references
    references.each do |sheet,cells|
      current_sheet << sheet
      cells.each do |ref,ast|
        count_dependencies_for(sheet,ref,ast)
      end
      current_sheet.pop
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
    operator = ast[0]
    if respond_to?(operator)
      send(operator,*ast[1..-1])
    else
      ast[1..-1].each do |a|
        map(a)
      end
    end
  end
  
  def sheet_reference(sheet,reference)
    ref = reference.last.gsub('$','')
    @dependencies[sheet][ref] ||= 0
    @dependencies[sheet][ref] += 1
  end
  
  def cell(reference)
    ref = reference.gsub('$','')
    @dependencies[current_sheet.last][ref] ||= 0
    @dependencies[current_sheet.last][ref] += 1
  end
   
end
