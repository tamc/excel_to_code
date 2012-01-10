class ReplaceTableReferenceAst
  
  attr_accessor :tables, :worksheet, :cell
  
  def initialize(tables, worksheet = nil, cell = nil)
    @tables, @worksheet, @cell = tables, worksheet, cell
  end
  
  def map(ast)
    if ast.is_a?(Array)
      operator = ast.shift
      if respond_to?(operator)
        send(operator,*ast)
      else
        [operator,*ast.map {|a| map(a) }]
      end
    else
      return ast
    end
  end
  
  def table_reference(table_name,table_reference)
    tables[table_name.downcase].reference_for(table_name,table_reference,worksheet,cell)
  end
  
end


class ReplaceTableReferences
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,sheet_name,table_data,output)
    tables = {}
    table_data.each do |line|
      table = Table.new(*line.strip.split("\t"))
      tables[table.name.downcase] = table
    end
        
    rewriter = ReplaceTableReferenceAst.new(tables,sheet_name)
  
    input.lines do |line|
      # Looks to match shared string lines
      if line =~ /\[:table_reference/
        ref, ast = line.split("\t")
        rewriter.cell = ref
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
  end
  
end