class ReplaceArithmeticOnRangesAst
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    ast.each do |a| 
      next unless a.is_a?(Array)
      case ast.first
      when :error, :null, :space, :prefeix, :boolean_true, :boolean_false, :number, :string
        next
      when :sheet_reference, :table_reference, :local_table_reference
        next
      else
        map(a)
      end
    end
    arithmetic(ast) if ast.first == :arithmetic
    comparison(ast) if ast.first == :comparison
    ast
  end

  # FIXME: DRY THIS UP
  def comparison(ast)
    left, operator, right = ast[1], ast[2], ast[3]
    # Three different options, array on the left, array on the right, or both
    # array on the left first
    if left.first == :array && right.first != :array
      ast.replace(
        array_map(left) do |cell|
          [:comparison, map(cell), operator, right]
        end
      )

    # array on the right next
    elsif left.first != :array && right.first == :array
      ast.replace(
        array_map(right) do |cell|
          [:comparison, left, operator, map(cell)]
        end
      )
    
    # now array both sides
    elsif left.first == :array && right.first == :array 
      ast.replace(
        left.map.with_index do |row, i|
          if row == :array
            row
          else
            row.map.with_index do |cell, j|
              if cell == :row
                cell
              elsif i >= left.length || i >= right.length || j >= left[1].length || j >= right[1].length
                [:error, "#VALUE!"]
              else
                [:comparison, map(left[i][j]), operator, map(right[i][j])]
              end
            end
          end
        end
      )
    end
  end

  
  # Format [:artithmetic, left, operator, right] 
  # should have removed arithmetic with more than one operator
  # in an earlier transformation
  def arithmetic(ast)
    left, operator, right = ast[1], ast[2], ast[3]
    # Three different options, array on the left, array on the right, or both
    # array on the left first
    if left.first == :array && right.first != :array
      ast.replace(
        array_map(left) do |cell|
          [:arithmetic, map(cell), operator, right]
        end
      )

    # array on the right next
    elsif left.first != :array && right.first == :array
      ast.replace(
        array_map(right) do |cell|
          [:arithmetic, left, operator, map(cell)]
        end
      )
    
    # now array both sides
    elsif left.first == :array && right.first == :array 
      ast.replace(
        left.map.with_index do |row, i|
          if row == :array
            row
          else
            row.map.with_index do |cell, j|
              if cell == :row
                cell
              elsif i >= left.length || i >= right.length || j >= left[1].length || j >= right[1].length
                [:error, "#VALUE!"]
              else
                [:arithmetic, map(left[i][j]), operator, map(right[i][j])]
              end
            end
          end
        end
      )
    end
  end

  def array_map(array)
    array.map do |row|
      if row == :array
        row
      else
        row.map do |column|
          if column == :row
            column
          else
            yield column
          end
        end
      end
    end
  end
    
end
