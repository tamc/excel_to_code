require_relative 'map_formulae_to_values'

class ReplaceFormulaeWithCalculatedValues
    
  attr_accessor :excel_file

  def self.replace(*args)
    self.new.replace(*args)
  end

  attr_accessor :replacements_made_in_the_last_pass
  
  def replace(input,output)
    rewriter = MapFormulaeToValues.new
    rewriter.original_excel_filename = excel_file
    input.each_line do |line|
      begin
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
        rewriter.reset
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end
    end
    @replacements_made_in_the_last_pass = rewriter.replacements_made_in_the_last_pass
  end
end
