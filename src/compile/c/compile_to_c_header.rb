class CompileToCHeader
  
  attr_accessor :worksheet
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,sheet_names_file,output,defaults = nil)
    c_name = Hash[sheet_names_file.readlines.map { |line| line.strip.split("\t")}][worksheet]
    input.lines do |line|
      begin
        ref, formula = line.split("\t")
        output.puts "ExcelValue #{c_name}_#{ref.downcase}();"
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end      
    end
  end  
end
