class IdentifyDependencies
  
  attr_accessor :references
  attr_accessor :dependencies
  attr_accessor :current_sheet
  attr_accessor :circular_reference_check
  
  def initialize(references = {}, dependencies = {})
    @references = references
    @dependencies = dependencies
    dependencies.default_proc = lambda do |hash,key|
      hash[key] = {}
    end
    @current_sheet = []
    @circular_reference_check = [] 
  end
  
  def add_depedencies_for(sheet,cell = :all)
    if cell == :all
      references.each do |ref,ast|
        next unless ref.first == sheet
        circular_reference_check.clear
        recursively_add_dependencies_for(sheet,ref.last)
      end
    else
      circular_reference_check.clear
      recursively_add_dependencies_for(sheet,cell)
    end
  end
    
  def recursively_add_dependencies_for(sheet,cell)
    return if circular_reference?([sheet, cell])
    return if dependencies[sheet].has_key?(cell)
    dependencies[sheet][cell] = true
    ast = references[[sheet,cell]]
    return unless ast
    current_sheet.push(sheet)
    begin
      map(ast)
    rescue ExcelToCodeException
      puts "[#{sheet}, #{cell}] => #{ast}"
      raise
    end
    current_sheet.pop
    circular_reference_check.pop
  end

  def circular_reference?(ref)
    if circular_reference_check.include?(ref)
      raise ExcelToCodeException.new("Possible circular reference in #{circular_reference_check} #{ref}")
    else
      circular_reference_check.push(ref)
    end
    false
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
    recursively_add_dependencies_for(sheet,Reference.for(reference.last).unfix.to_sym)
  end
  
  def cell(reference)
    recursively_add_dependencies_for(current_sheet.last,Reference.for(reference).unfix.to_sym)
  end
   
end
