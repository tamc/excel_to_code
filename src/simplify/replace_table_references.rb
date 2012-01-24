class ReplaceTableReferenceAst
  
  attr_accessor :tables, :worksheet, :referring_cell
  
  def initialize(tables, worksheet = nil, referring_cell = nil)
    @tables, @worksheet, @referring_cell = tables, worksheet, referring_cell
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
  
  def table_reference(table_name,table_reference)
    tables[table_name.downcase].reference_for(table_name,table_reference,worksheet,referring_cell)
  end
  
end


class ReplaceTableReferences
  
  attr_accessor :sheet_name
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,table_data,output)
    tables = {}
    table_data.each do |line|
      table = Table.new(*line.strip.split("\t"))
      tables[table.name.downcase] = table
    end
        
    rewriter = ReplaceTableReferenceAst.new(tables,sheet_name)
  
    input.lines do |line|
      # Looks to match shared string lines
      if line =~ /\[:table_reference/
        cols = line.split("\t")
        ast = cols.pop
        ref = cols.first
        rewriter.referring_cell = ref
        output.puts "#{cols.join("\t")}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
  end
  
end