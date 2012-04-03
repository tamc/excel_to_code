class CompileToCHeader
  
  attr_accessor :worksheet
  attr_accessor :gettable
  attr_accessor :settable
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,sheet_names_file,output,defaults = nil)
    c_name = Hash[sheet_names_file.readlines.map { |line| line.strip.split("\t")}][worksheet]
    self.gettable ||= lambda { |ref| true }
    self.settable ||= lambda { |ref| false }
    input.lines do |line|
      begin
        ref, formula = line.split("\t")
        static_or_not = (gettable.call(ref) || settable.call(ref)) ? "" : "static "
        output.puts "#{static_or_not}ExcelValue #{c_name}_#{ref.downcase}();"
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end      
    end
  end  
end
