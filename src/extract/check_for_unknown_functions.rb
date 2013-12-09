require_relative '../compile/ruby/map_formulae_to_ruby'

class CheckForUnknownFunctions
  
  attr_accessor :settable
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def check(input,output)
    self.settable ||= lambda { |ref| false }
    input.each_line do |line|
      line.scan(/\[:function, :["']?([A-Z ]+)['"]?/).each do |match|
        output.puts $1 unless MapFormulaeToRuby::FUNCTIONS.has_key?($1.to_sym)
      end
    end
  end
  
end
