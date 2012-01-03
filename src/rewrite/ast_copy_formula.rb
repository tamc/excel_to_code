require_relative '../excel'

class AstCopyFormula
  attr_accessor :rows_to_move
  attr_accessor :columns_to_move
  
  def initialize
    self.rows_to_move = 0
    self.columns_to_move = 0
  end
  
  def copy(ast)
    return ast unless ast.is_a?(Array)
    operator = ast.shift
    if respond_to?(operator)
      send(operator,*ast)
    else
      [operator,*ast.map {|a| copy(a) }]
    end
  end
  
  def cell(reference)
    r = Reference.for(reference)
    [:cell,r.offset(rows_to_move,columns_to_move)]
  end
  
  def area(reference)
    a = Area.for(reference)
    [:area,a.offset(rows_to_move,columns_to_move)]
  end
    
end