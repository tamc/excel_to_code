require_relative '../excel'
require_relative '../util/not_supported_exception'

class AstCopyFormula
  attr_accessor :rows_to_move
  attr_accessor :columns_to_move
  
  def initialize
    self.rows_to_move = 0
    self.columns_to_move = 0
  end
  
  def copy(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if respond_to?(operator)
      send(operator,*ast[1..-1])
    else
      [operator,*ast[1..-1].map {|a| copy(a) }]
    end
  end
  
  def cell(reference)
    r = Reference.for(reference)
    [:cell,r.offset(rows_to_move,columns_to_move)]
  end
  
  def area(start,finish)
    s = Reference.for(start).offset(rows_to_move,columns_to_move)
    f = Reference.for(finish).offset(rows_to_move,columns_to_move)
    [:area,s,f]
  end
    
  def column_range(reference)
    raise NotSupportedException.new("Column ranges not suported in AstCopyFormula")
  end

  def row_range(reference)
    raise NotSupportedException.new("Row ranges not suported in AstCopyFormula")
  end  
  
end