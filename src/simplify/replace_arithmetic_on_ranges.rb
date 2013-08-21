class ReplaceArithmeticOnRangesAst
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if respond_to?(operator)
      send(operator,*ast[1..-1])
    else
      [operator,*ast[1..-1].map {|a| map(a) }]
    end
  end
  
  def arithmetic(left, operator, right)
    if left.first == :array && right.first != :array
      mapped_right = map(right)
      array_map(left) do |cell|
        [:arithmetic, map(cell), operator, mapped_right]
      end
    elsif left.first != :array && right.first == :array
      mapped_left = map(left)
      array_map(right) do |cell|
        [:arithmetic, mapped_left, operator, map(cell)]
      end
    elsif left.first == :array && right.first == :array 
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
    else
      [:arithmetic, map(left), operator, map(right)]
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
