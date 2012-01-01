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
  
  def calculate_excel_variables
    return if @excel_variables_calculated
    self =~ /([^:]+):(.*)/
    @excel_start = Reference.for($1)
    @excel_finish = Reference.for($2)
  end
  
  def offset(row,column)
    calculate_excel_variables
    Area.for([
      @excel_start.offset(row,column),
      ':',
      @excel_finish.offset(row,column),
    ].join)
  end

end