require_relative 'reference'

class Area < String
  
  # This is so that we only have one instance for a given area
  @@areas_for_text ||= Hash.new do |hash,text|
    hash[text] = area = Area.new(text)
    area
  end
    
  # This is so that we only have one instance of a given reference specified by its variables
  def Area.for(text)
     @@areas_for_text[text]
  end
  
  attr_reader :excel_start, :excel_finish
  
  def calculate_excel_variables
    return if @excel_variables_calculated
    self =~ /([^:]+):(.*)/
    @excel_start = Reference.for($1)
    @excel_finish = Reference.for($2)
    @excel_start.calculate_excel_variables
    @excel_finish.calculate_excel_variables
  end
  
  def offset(row,column)
    calculate_excel_variables
    Area.for([
      @excel_start.offset(row,column),
      ':',
      @excel_finish.offset(row,column),
    ].join)
  end
  
  def offsets
    calculate_excel_variables
    
    columns = @excel_finish.excel_column_number -  @excel_start.excel_column_number
    rows = @excel_finish.excel_row_number -  @excel_start.excel_row_number
    Enumerator.new do |yielder|
      0.upto(columns).each do |c|
        0.upto(rows).each do |r|
          yielder.yield([c,r])
        end
      end
    end
  end

end