class MapValuesToConstants
  
  attr_accessor :constants
  
  def initialize
    count = 0
    @constants = Hash.new do |hash,key|
      count += 1
      hash[key] = "constant#{count}"
    end
  end
  
  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if replace?(ast)
      ast.replace([:constant, constants[ast.dup]])
    else
      ast.each { |a| map(a) }
    end
  end

  def replace?(ast)
    case ast.first
    when :string; return true
    when :percentage, :number
      n = ast.last.to_f
      # Don't use constant if an integer less than ten, 
      # as will use constant predefined in the c runtime
      return true if n > 10
      return true if n < 0
      return false if n % 1 == 0
      return true
    else
      return false
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
