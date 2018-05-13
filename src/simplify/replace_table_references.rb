class ReplaceTableReferenceAst

  attr_accessor :tables, :worksheet, :referring_cell, :function_needs_area

  def initialize(tables, worksheet = nil, referring_cell = nil)
    @tables, @worksheet, @referring_cell = tables, worksheet, referring_cell
    @function_needs_area = false
  end

  def map(ast)
    return ast unless ast.is_a?(Array)
    case ast[0]
    when :function;
      check_function_needs_area(ast)
    when :table_reference;
      table_reference(ast)
    when :local_table_reference;
      local_table_reference(ast)
    else
      ast.each { |a| map(a) }
    end
    ast
  end

  def check_function_needs_area(ast)
    if ast[1] == :MATCH
      @function_needs_area = false
      map(ast[2])
      @function_needs_area = true
      map(ast[3])
      @function_needs_area = false
      ast[4..-1].each { |a| map(a) }
    else
      @function_needs_area = false
      ast.each { |a| map(a) }
    end
    ast
  end

  # Of the format [:table_reference, table_name, table_reference]
  def table_reference(ast)
    table_name = ast[1]
    table_reference = ast[2]
    return ast.replace([:error, :"#REF!"]) unless tables.has_key?(table_name.downcase)
    ast.replace(tables[table_name.downcase].reference_for(table_name,table_reference,worksheet,referring_cell, @function_needs_area))
  end

  # Of the format [:local_table_reference, table_reference]
  def local_table_reference(ast)
    table_reference = ast[1]
    table = tables.values.find do |table|
      table.includes?(worksheet,referring_cell)
    end
    return ast.replace([:error, :"#REF!"]) unless table
    ast.replace(table.reference_for(table.name,table_reference,worksheet,referring_cell, @function_needs_area))
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

    input.each_line do |line|
      # Looks to match shared string lines
      begin
        if line =~ /\[(:table_reference|:local_table_reference)/
          cols = line.split("\t")
          ast = cols.pop
          ref = cols.first
          rewriter.referring_cell = ref
          output.puts "#{cols.join("\t")}\t#{rewriter.map(eval(ast)).inspect}"
        else
          output.puts line
        end
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end
    end
  end

end
