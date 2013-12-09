require 'nokogiri'

class ExtractValues < Nokogiri::XML::SAX::Document 

  def self.extract(*args)
    self.new.extract(*args)
  end

  def extract(sheet_name, input)
    @sheet_name = sheet_name.to_sym
    @output = {}
    @parsing = false
    Nokogiri::XML::SAX::Parser.new(self).parse(input)
    @output
  end
    
  attr_accessor :ref, :type, :value
  
  def start_element(name,attributes)
    if name == "v"
      @parsing = true 
      @value = []
    elsif name == "c"
      @ref = attributes.assoc('r').last.to_sym
      type = attributes.assoc('t')
      @type = type ? type.last : "n"
    end
  end
  
  def end_element(name)
    return unless name == "v"
    @parsing = false
    value = @value.join
    ast = case @type
    when 'b'; value == "1" ? [:boolean_true] : [:boolean_false]
    when 's'; [:shared_string,value.to_i]
    when 'n'; [:number,value.to_f]
    when 'e'; [:error,value.to_sym]
    when 'str'; [:string,value.gsub(/_x[0-9A-F]{4}_/,'').freeze]
    else
      $stderr.puts "Type #{type} not known #{@sheet_name} #{@ref}"
      exit
    end
    @output[[@sheet_name, @ref]] = ast
  end
  
  def characters(string)
    return unless @parsing
    @value << string
  end
  
end
