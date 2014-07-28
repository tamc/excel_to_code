require 'ox'

class ExtractTable < ::Ox::Sax

    def self.extract(*args)
      self.new.extract(*args)
    end

    def extract(sheet_name, input_xml)
      @sheet_name = sheet_name
      @output = {} 
      @table_name = nil
      @table_ref = nil
      @table_total_rows = nil
      @table_columns = nil

      @state = :not_parsing
      Ox.sax_parse(self, input_xml, :convert_special => true)
      @output
    end
    
  def start_element(name)
    case name
    when :table
      @table_name = nil
      @table_ref = nil
      @table_total_rows = nil
      @table_columns = []
      @state = :parsing_table
    when :tableColumn
      @state = :parsing_table_column
    else
      @state = :not_parsing
    end
  end

  def attr(name, value)
    case @state
    when :not_parsing
      return
    when :parsing_table_column
      case name
      when :name
        @table_columns << value
      end
    when :parsing_table
      case name
      when :displayName
        @table_name = value
      when :ref
        @table_ref = value
      when :totalsRowCount
        @table_total_rows = value
      end
    end
  end
  
  def end_element(name)
    case name
    when :tableColumn
      @state = :parsing_table
    when :table
      @output[@table_name] = [@sheet_name, @table_ref, @table_total_rows || "0", *@table_columns]
      @state = :not_parsing
    end

  end  
end
