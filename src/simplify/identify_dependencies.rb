class IdentifyDependencies
  
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
  
  def add_depedencies_for(sheet,cell = :all)
    if cell == :all
      references.each do |ref,ast|
        next unless ref.first == sheet
        recursively_add_dependencies_for(sheet,ref.last)
      end
    else
      recursively_add_dependencies_for(sheet,cell)
    end
  end
    
  def recursively_add_dependencies_for(sheet,cell)
    return if dependencies[sheet].has_key?(cell)
    dependencies[sheet][cell] = true
    ast = references[[sheet,cell]]
    return unless ast
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
      ast.each {|a| map(a) }
    end
  end
  
  def sheet_reference(sheet,reference)
    recursively_add_dependencies_for(sheet,reference.last.gsub('$',''))
  end
  
  def cell(reference)
    recursively_add_dependencies_for(current_sheet.last,reference.gsub('$',''))
  end
   
end
