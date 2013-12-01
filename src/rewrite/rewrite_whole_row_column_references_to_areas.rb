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
  
end

class MapColumnAndRowRangeAst 
  
  attr_accessor :default_worksheet_name, :worksheet_dimensions
  
  def initialize(default_worksheet_name,worksheet_dimensions)
    @default_worksheet_name, @worksheet_dimensions = default_worksheet_name, worksheet_dimensions
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
  
  def sheet_reference(sheet_name,reference)
    if reference.first == :row_range
      [:sheet_reference,sheet_name,[:area,*worksheet_dimensions[sheet_name].map_row(reference[1],reference[2])]]
    elsif reference.first == :column_range
      [:sheet_reference,sheet_name,[:area,*worksheet_dimensions[sheet_name].map_column(reference[1],reference[2])]]
    else
      [:sheet_reference,sheet_name,reference]
    end
  end
  
  def row_range(start,finish)
    [:area,*worksheet_dimensions[default_worksheet_name].map_row(start,finish)]
  end
  
  def column_range(start,finish)
    [:area,*worksheet_dimensions[default_worksheet_name].map_column(start,finish)]
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
