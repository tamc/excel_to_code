require_relative 'reference'

class Area < String
  
  # This is so that we only have one instance for a given area
  @@areas_for_text ||= Hash.new do |hash,text|
    hash[text] = area = Area.new(text)
    area
  end
    
  # This is so that we only have one instance of a given reference specified by its variables
  def Area.for(text)
     @@areas_for_text[text.to_s]
  end
  
  attr_reader :excel_start, :excel_finish
  
  def calculate_excel_variables
    return if @excel_variables_calculated
    if self =~ /([^:]+):(.*)/
      @excel_start = Reference.for($1)
      @excel_finish = Reference.for($2)
    else
      @excel_start = @excel_finish = Reference.for(self)
    end      
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
    Enumerator.new do |yielder|
      0.upto(width).each do |c|
        0.upto(height).each do |r|
          yielder.yield([r,c])
        end
      end
    end
  end
  
  def height
    calculate_excel_variables
    @excel_finish.excel_row_number -  @excel_start.excel_row_number
  end
  
  def width
    calculate_excel_variables
    @excel_finish.excel_column_number -  @excel_start.excel_column_number
  end
  
  def to_array_literal(sheet = nil)
    calculate_excel_variables
    unfixed_start = @excel_start.unfix
    fc = CachingFormulaParser.instance
    [:array,
      *(0.upto(height).map do |row|
        [:row,
          *(0.upto(width).map do |column|
            if sheet
              fc.sheet_reference(
                [:sheet_reference, sheet, [:cell, unfixed_start.offset(row,column)]]
              )
            else
              [:cell,
                unfixed_start.offset(row,column)
              ]
            end
          end)
        ]
      end)
    ]
  end
  
  def includes?(reference)
    calculate_excel_variables
    r = Reference.for(reference)
    r.calculate_excel_variables
    return false if r.excel_row_number < @excel_start.excel_row_number || r.excel_row_number > @excel_finish.excel_row_number
    return false if r.excel_column_number < @excel_start.excel_column_number || r.excel_column_number > @excel_finish.excel_column_number
    true
  end
  
end
