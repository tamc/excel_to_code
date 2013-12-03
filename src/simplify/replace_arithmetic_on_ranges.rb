class ReplaceArithmeticOnRangesAst
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    arithmetic(ast) if ast.first == :arithmetic
    ast.each { |a| map(a) }
    ast
  end
  
  # Format [:artithmetic, left, operator, right] 
  # should have removed arithmetic with more than one operator
  # in an earlier transformation
  def arithmetic(ast)
    left, operator, right = ast[1], ast[2], ast[3]
    # Three different options, array on the left, array on the right, or both
    # array on the left first
    if left.first == :array && right.first != :array
      map(right)
      ast.replace(
        array_map(left) do |cell|
          [:arithmetic, map(cell), operator, right]
        end
      )

    # array on the right next
    elsif left.first != :array && right.first == :array
      map(left)
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
              elsif i >= left.length || i >= right.length || j >= left.first.length || j >= right.first.length
                [:error, "#VALUE!"]
              else
                [:arithmetic, map(left[i][j]), operator, map(right[i][j])]
              end
            end
          end
        end
      )
    else
      map(left)
      map(right)
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
  

class ReplaceArithmeticOnRanges
    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    rewriter = ReplaceArithmeticOnRangesAst.new
    input.each_line do |line|
      # Looks to match lines with references
      if line =~ /:arithmetic/ && line =~ /:array/
        content = line.split("\t")
        ast = eval(content.pop)
        output.puts "#{content.join("\t")}\t#{rewriter.map(ast).inspect}"
      else
        output.puts line
      end
    end
  end
end
