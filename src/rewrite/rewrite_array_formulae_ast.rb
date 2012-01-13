require_relative '../excel'

class RewriteArrayFormulaAst
  
  attr_accessor :row_offset, :column_offset
  
  def initialize
    @row_offset = 0
    @column_offset = 0
    @offsetting = [true]
  end
  
  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if respond_to?(operator)
      send(operator,*ast[1..-1])
    else
      [operator,*ast[1..-1].map {|a| map(a) }]
    end
  end
  
  def area(start,finish)
    if @offsetting.last
      c = Reference.for(start).unfix
      [:cell,c.offset(row_offset,column_offset)]
    else
      [:area,start,finish]
    end
  end
  
  def function(name,*arguments)
    if respond_to?("map_#{name.downcase}")
      send("map_#{name.downcase}",*arguments)
    else
      [:function,name,*arguments.map { |a| map(a)}]
    end
  end
  
  def map_sum(*arguments)
    @offsetting << false
    a = [:function,"SUM",*arguments.map { |a| map(a)}]
    @offsetting.pop
    return a
  end
  
end