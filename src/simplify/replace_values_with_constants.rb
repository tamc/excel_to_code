class MapValuesToConstants
  
  attr_accessor :constants
  
  def initialize
    count = 0
    @constants = Hash.new do |hash,key|
      count += 1
      hash[key] = "C#{count}"
    end
  end
  
  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if [:number,:percentage,:string].include?(operator)
      ast.replace([:constant, constants[ast]])
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
