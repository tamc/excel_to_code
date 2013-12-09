class MapNamedReferenceToCSetter

  attr_accessor :sheet_names
  attr_accessor :cells_that_can_be_set_at_runtime

  def initialize
    reset
  end

  def reset
    @new_value_name = "newValue"
  end

  def map(ast)
    if ast.is_a?(Array)
      operator = ast[0]
      if respond_to?(operator)
        send(operator,*ast[1..-1])
      else
        raise NotSupportedException.new("#{operator} in #{ast.inspect} not supported")
      end
    else
      raise NotSupportedException.new("#{ast} not supported")
    end
  end

  def cell(reference)
    Reference.for(reference).unfix.downcase.to_s
  end

  def sheet_reference(sheet,reference)
    s = sheet_names[sheet]
    c = map(reference)
    return "  // #{s}_#{c} not settable" unless settable(sheet, c)
    "  set_#{s}_#{c}(#{@new_value_name});"
  end

  def array(*rows)
    counter = -1

    result = rows.map do |r|
      r.shift if r.first == :row
      r.map do |c| 
        counter += 1
        @new_value_name = "array[#{counter}]"
        map(c)
      end
    end.flatten.join("\n")
    
    @new_value_name = "newValue"

    "  ExcelValue *array = newValue.array;\n#{result}"
  end

  def settable(sheet, reference)
    settable_refs = cells_that_can_be_set_at_runtime[sheet]
    return false unless settable_refs
    return true if settable_refs == :all
    settable_refs.include?(reference.upcase) 
  end

end

class CompileNamedReferenceSetters

  attr_accessor :cells_that_can_be_set_at_runtime

  def self.rewrite(*args)
    new.rewrite(*args)
  end

  def rewrite(named_references, sheet_names, output)
    mapper = MapNamedReferenceToCSetter.new
    mapper.sheet_names = sheet_names
    mapper.cells_that_can_be_set_at_runtime = cells_that_can_be_set_at_runtime

    named_references.each do |name, ast|
      output.puts "void set_#{name}(ExcelValue newValue) {"
      output.puts mapper.map(ast)
      output.puts "}"
      output.puts
    end
    output
  end



end
