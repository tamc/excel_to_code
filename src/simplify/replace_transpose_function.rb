class ReplaceTransposeFunction

  def replace(ast)
    map(ast)
  end
   
  def map(ast)
    return ast unless ast.is_a?(Array)
    function(ast) if ast[0] == :function
    ast.each { |a| map(a) }
    ast
  end
  
  def function(ast)
    return unless ast[1] == :TRANSPOSE
    array = array_ast_to_ruby_array(ast[2])
    array = array.transpose
    ast.replace(ruby_array_to_array_ast(array))
    ast
  end

  def array_ast_to_ruby_array(array_ast)
    return [[array_ast]] unless array_ast.first == :array
    array_ast[1..-1].map do |row_ast|
      row_ast[1..-1].map do |cell|
        cell
      end
    end
  end

  def ruby_array_to_array_ast(ruby_array)
    [:array].concat(ruby_array.map { |row| [:row].concat(row) })
  end
end
