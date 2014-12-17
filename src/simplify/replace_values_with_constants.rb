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

