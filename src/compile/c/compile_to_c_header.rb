class CompileToCHeader
  
  attr_accessor :gettable
  attr_accessor :settable
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(formulae, c_name_for_worksheet_name, output)
    self.gettable ||= lambda { |ref| true }
    self.settable ||= lambda { |ref| false }
    formulae.each do |ref, ast|
      begin
        static_or_not = (gettable.call(ref) || settable.call(ref)) ? "" : "static "
        worksheet = c_name_for_worksheet_name[ref.first]
        ref = ref.last.downcase
        output.puts "#{static_or_not}ExcelValue #{worksheet}_#{ref}();"
      rescue Exception => e
        puts "Exception at  #{ref} #{ast}"
        raise
      end      
    end
  end  
end
