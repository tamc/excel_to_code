require_relative '../excel'

class ExpandArrayFormulaeAst
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if respond_to?(operator)
      send(operator,*ast[1..-1])
    else
      [operator,*ast[1..-1].map {|a| map(a) }]
    end
  end

  def arithmetic(left,operator,right)
    left = map(left)
    right = map(right)
    return [:arithmetic, left, operator, right] unless array?(left,right)
    
    a = array_ast_to_ruby_array(left)
    b = array_ast_to_ruby_array(right)
    
    a = Array.new(b.length,Array.new(b.first.length,a)) unless a.is_a?(Array)
    b = Array.new(a.length,Array.new(a.first.length,b)) unless b.is_a?(Array)
    
    return [:error, "#VALUE!"] unless b.length == a.length
    return [:error, "#VALUE!"] unless b.first.length == a.first.length
    
    [:array, *a.map.with_index do |row,i|
      [:row, *row.map.with_index do |cell,j|
        [:arithmetic, cell, operator, b[i][j]]
      end ]      
    end]
  end

  def comparison(left,operator,right)
    left = map(left)
    right = map(right)
    return [:comparison, left, operator, right] unless array?(left,right)
    
    a = array_ast_to_ruby_array(left)
    b = array_ast_to_ruby_array(right)
    
    a = Array.new(b.length,Array.new(b.first.length,a)) unless a.is_a?(Array)
    b = Array.new(a.length,Array.new(a.first.length,b)) unless b.is_a?(Array)
    
    return [:error, "#VALUE!"] unless b.length == a.length
    return [:error, "#VALUE!"] unless b.first.length == a.first.length
    
    [:array, *a.map.with_index do |row,i|
      [:row, *row.map.with_index do |cell,j|
        [:comparison, cell, operator, b[i][j]]
      end ]      
    end]
  end
  
  def function(name,*arguments)
    if respond_to?("map_#{name.downcase}")
      send("map_#{name.downcase}",*arguments)
    else
      [:function,name,*arguments.map { |a| map(a)}]
    end
  end
  
  def map_sum(*arguments)
    [:function,"SUM",*arguments.map { |a| map(a)}]
  end
  
  def map_cosh(argument)
    a = map(argument)
    return [:function, "COSH", a] unless array?(a)
    [:array, *array_ast_to_ruby_array(a).map do |row|
      [:row, *row.map do |cell|
        [:function, "COSH", cell]
      end ]
    end ]
  end
  
  
  private
  
  def array?(*args)
    args.any? { |a| a.first == :array }
  end
  
  def array_ast_to_ruby_array(array_ast)
    return array_ast unless array_ast.first == :array
    array_ast[1..-1].map do |row_ast|
      row_ast[1..-1].map do |cell|
        cell
      end
    end
  end
  
end
