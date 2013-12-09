require_relative '../excel'

class WorksheetDimension
  
  attr_reader :first_column, :last_column, :first_row, :last_row
  
  def initialize(area)
    @area = Area.for(area)
    @area.calculate_excel_variables
    @first_column = @area.excel_start.excel_column
    @last_column =  @area.excel_finish.excel_column
    @first_row = @area.excel_start.excel_row
    @last_row = @area.excel_finish.excel_row
  end
  
  def map_row(start,finish)
    ["#{first_column}#{start}","#{last_column}#{finish}"]
  end

  def map_column(start,finish)
    ["#{start}#{first_row}","#{finish}#{last_row}"]
  end

  def inspect
    area.inspect
  end
  
end

class MapColumnAndRowRangeAst 
  
  attr_accessor :default_worksheet_name, :worksheet_dimensions
  
  def initialize(default_worksheet_name,worksheet_dimensions)
    @default_worksheet_name, @worksheet_dimensions = default_worksheet_name, worksheet_dimensions
    @worksheet_names = []
  end
  
  def map(ast)
    return ast unless ast.is_a?(Array)
    case ast[0]
    when :sheet_reference; sheet_reference(ast)
    when :row_range; row_range(ast)
    when :column_range; column_range(ast)
    end
    ast.each {|e| map(e)}
  end
  
  # Of the form [:sheet_reference, sheet_name, reference]
  def sheet_reference(ast)
    @worksheet_names.push(ast[1].to_sym) #FIXME: Remove once all symbols
    map(ast[2])
    @worksheet_names.pop
  end
  
  # Of the form [:row_range, start, finish]
  def row_range(ast)
    worksheet = @worksheet_names.last || @default_worksheet_name
    ast.replace([:area,*worksheet_dimensions[worksheet].map_row(ast[1],ast[2])])
  end
  
  # Of the form [:column_range, start, finish]
  def column_range(ast)
    worksheet = @worksheet_names.last || @default_worksheet_name
    ast.replace([:area,*worksheet_dimensions[worksheet].map_column(ast[1],ast[2])])
  end
  
end

class RewriteWholeRowColumnReferencesToAreas
  
  attr_accessor :sheet_name
  attr_accessor :dimensions
  
  def self.rewrite(input,output)
    new.rewrite(input,output)
  end
  
  def rewrite(input,output)
    input.each_line do |line|
      if line =~ /(:column_range|:row_range)/
        content = line.split("\t")
        ast = eval(content.pop)
        output.puts "#{content.join("\t")}\t#{mapper.map(ast).inspect}"
      else
        output.puts line
      end
    end
  end
  
  def worksheet_dimensions=(worksheet_dimensions)
    @dimensions = {}
    worksheet_dimensions.each do |name, area|
      @dimensions[name] = WorksheetDimension.new(area)
    end
    @mapper = nil
  end
  
  def mapper
    @mapper ||= MapColumnAndRowRangeAst.new(sheet_name,dimensions)
  end
  
end
