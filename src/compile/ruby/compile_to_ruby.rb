require_relative 'map_formulae_to_ruby'

class CompileToRuby
  
  attr_accessor :settable
  attr_accessor :worksheet
  attr_accessor :defaults
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input, sheet_names, output)
    self.settable ||= lambda { |ref| false }
    self.defaults ||= []
    mapper = MapFormulaeToRuby.new
    mapper.sheet_names = sheet_names
    input.each do |ref, ast|
      begin
        worksheet = ref.first.to_s
        cell = ref.last
        mapper.worksheet = worksheet
        c_name = mapper.sheet_names[worksheet]
        name = c_name ? "#{c_name}_#{cell.downcase}" : cell.downcase
        if settable.call(ref)
          output.puts "  attr_accessor :#{name} # Default: #{mapper.map(ast)}"
          defaults << "    @#{name} = #{mapper.map(ast)}"
        else
          output.puts "  def #{name}; @#{name} ||= #{mapper.map(ast)}; end"
        end
      rescue Exception => e
        puts "Exception at #{ref} => #{ast}"
        raise
      end      
    end
  end
  
end
