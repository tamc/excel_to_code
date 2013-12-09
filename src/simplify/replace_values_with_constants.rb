class MapValuesToConstants
  
  attr_accessor :constants
  
  def initialize
    count = 0
    @constants = Hash.new do |hash,key|
      count += 1
      hash[key] = "constant#{count}"
    end
  end
  
  POTENTIAL_CONSTANTS = { :number => true, :percentage => true, :string => true}

  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if POTENTIAL_CONSTANTS.has_key?(operator)
      ast.replace([:constant, constants[ast.dup]])
    else
      ast.each { |a| map(a) }
    end
  end

end


class ReplaceValuesWithConstants
  
  attr_accessor :rewriter
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    @rewriter ||= MapValuesToConstants.new
    input.each_line do |line|
      begin
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end
    end

  end
end
