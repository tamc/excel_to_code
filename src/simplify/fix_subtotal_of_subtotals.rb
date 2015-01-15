class FixSubtotalOfSubtotals

  attr_accessor :references
  attr_accessor :count_replaced
  
  def initialize(references = nil)
    @references = references
    @count_replaced = 0
  end

  def map(ast)
    return ast unless ast.is_a?(Array)
    if ast[0] == :function && ast[1] == :SUBTOTAL
      count_before = @count_replaced
      new_arguments = ast[2..-1].map do |a| 
        remove_subtotals_from(a)
      end
      new_arguments.compact!
      if @count_replaced > count_before
        ast.replace([:function, :SUBTOTAL].concat(new_arguments))
      end
    end
    ast.each do |a|
      map(a) if a.is_a?(Array)
    end
  end

  def is_subtotal?(ast)
    return false unless ast.is_a?(Array)
    return true if ast[0] == :function && ast[1] == :SUBTOTAL
    ast.each do |e|
      r = is_subtotal?(e) if e.is_a?(Array) # To limit the stack depth
      return true if r
    end
    return false
  end

  def is_or_refers_to_subtotal?(ast)
    return false unless ast.is_a?(Array)
    return is_subtotal?(ast) unless ast[0] == :sheet_reference

    raise ExcelToCodeException.new("Expecting cell") unless ast[2][0] == :cell
    reference_ast = references[[ast[1], ast[2][1]]]
    return false unless reference_ast
    
    is_subtotal?(reference_ast)
  end

  def remove_subtotals_from(ast)
    return ast unless ast.is_a?(Array)
    if ast.first == :array
      new_ast = ast.dup
      new_ast.delete_if do |element|
        if element == :array
          false
        else
          raise ExcelToCodeException("Expecting row") unless element.is_a?(Array) && element[0] == :row
          element[1..-1].any? { |cell| is_or_refers_to_subtotal?(cell) }
        end
      end
      return ast if new_ast.length == ast.length
      @count_replaced += 1
      return new_ast
    else
      if is_or_refers_to_subtotal?(ast)
        @count_replaced += 1
        return nil 
      end
    end
    ast
  end
  



end
