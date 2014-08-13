# coding: utf-8
require_relative 'excel_to_x'
require 'ffi'

class ExcelToCFunction < ExcelToC
    
  def write_out_excel_as_code
    log.info "Writing C code"
        
    number_of_refs = @formulae.size + @named_references_to_keep.size
        
    # Output the workbook preamble
    o = output("#{output_name.downcase}.c")
    o.puts "// #{excel_file} approximately translated into C"
    o.puts "// definitions"
    o.puts "#define NUMBER_OF_REFS #{number_of_refs}"
    o.puts "#define EXCEL_FILENAME  #{excel_file.inspect}"
    o.puts "// end of definitions"
    o.puts

    o.puts '// First we have c versions of all the excel functions that we know'
    o.puts IO.readlines(File.join(File.dirname(__FILE__),'..','compile','c','excel_to_c_runtime.c')).join
    o.puts '// End of the generic c functions'
    o.puts 
    o.puts '// Start of the file specific functions'
    o.puts
    
    arguments = @named_references_that_can_be_set_at_runtime.map do |ref|
      "ExcelValue #{c_name_for(ref)}"
    end

    o.puts "void * run(#{arguments.join(',')}) {"
    o.puts
    o.puts "  // Transfer arguments into local variables"

    formulae_to_include = @formulae.dup

    @named_references_that_can_be_set_at_runtime.each do |ref|
      ast = @named_references[ref]
      case ast[0]
      when :array
        cols = ast[1].length-1
        ast[1..-1].map.with_index do |row,j|
          row[1..-1].map.with_index do |cell,i|
            raise Exception.new("Named reference #{name} contains #{ast} which isn't sheet_references") unless cell.first == :sheet_reference
            o.puts "  ExcelValue #{c_name_for_worksheet_name(cell[1])}_#{Reference.for(cell[2][1]).unfix.downcase.to_s} = ((ExcelValue*) #{c_name_for(ref)}.array)[#{(j*cols)+i}];"
            formulae_to_include.delete([cell[1], cell[2][1]])
          end
        end
      when :sheet_reference
        o.puts "  ExcelValue #{c_name_for_worksheet_name(ast[1])}_#{Reference.for(ast[2][1]).unfix.downcase.to_s} = #{c_name_for(ref)};"
        formulae_to_include.delete([ast[1], ast[2][1]])
      else
        p ref, ast
        exit
      end
    end
    variable_set_counter = 0
    
    o.puts
    o.puts "  // Set up the constants"
    mapper = MapValuesToCStructs.new
    @constants.each do |ref, ast|
      begin
        calculation = mapper.map(ast)
        o.puts "  const ExcelValue #{ref} = #{calculation};"
      rescue Exception => e
        puts "Exception at #{ref} #{ast}"
        raise
      end
    end          


    o.puts
    o.puts "  // Start doing the calculations"

    formulae_order = SortIntoCalculationOrder.new.sort(formulae_to_include)

    c = CompileToCFunction.new
    c.variable_set_counter = variable_set_counter
    # Output the elements from each worksheet in turn
    c.rewrite(formulae_to_include, @worksheet_c_names, o, formulae_order)

    m = MapFormulaeToCFunction.new
    m.counter = c.variable_set_counter

    m.sheet_names = @worksheet_c_names

    o.puts
    o.puts "  // Preparing results to return"
    o.puts "  ExcelValue *result_array = new_excel_value_array(#{@named_references_to_keep.size});"
    results = []
    @named_references_to_keep.each.with_index do |ref, i|
      ast = @named_references[ref] || @table_areas[ref]
      results << "  result_array[#{i}] = #{m.map(ast)}; // #{ref}"
    end
    o.puts "  "+m.initializers.join("\n  ")
    o.puts results.join("\n")
    o.puts "  return result_array;"



    o.puts "}; // End of run"


    close(o)
  end

  def default_value_for_named_reference(name)
    ast = @named_references[name]
    begin
      case ast.first
      when :sheet_reference
        # FIXME: EVAL!
        eval(MapValuesToRuby.new.map(@values[[ast[1], ast[2][1]]]))
      when :array
        ast[1..-1].map do |row|
          row[1..-1].map do |cell|
            raise Exception.new("Named reference #{name} contains #{ast} which isn't sheet_references") unless cell.first == :sheet_reference
            value_ast = @values[[cell[1], cell[2][1]]]
            if value_ast
              eval(MapValuesToRuby.new.map(value_ast))
            else
              nil 
            end
          end
        end
      end
    rescue Exception => e
      puts "Exception when finding default value for named reference #{name} => #{ast}"
      raise
    end
  end

  def write_fuby_ffi_interface
    log.info "Writing ruby FFI code"

    default_inputs = []
    input_names = {}
    output_names = []

    @named_references_that_can_be_set_at_runtime.each.with_index do |ref,i|
      default_inputs[i] = default_value_for_named_reference(ref) 
      input_names[c_name_for(ref)] = i
    end

    @named_references_to_keep.each.with_index do |ref, i|
      output_names[i] = c_name_for(ref)
    end

    name = output_name.downcase
    o = output("#{name}.rb")
      
    code = <<END
require 'ffi'
require 'singleton'

class #{ruby_module_name}

  INPUT_NAMES = #{input_names.inspect}
  OUTPUT_NAMES = #{output_names.inspect}

  def initialize
    reset
  end

  def reset
    @need_to_recalculate = true
    @inputs = #{default_inputs.inspect} # Defaults
    @outputs = {}
    C.reset
  end

  def method_missing(name, *arguments)
    name = name.to_s
    if arguments.size == 0
      get(name)
    elsif arguments.size == 1
      set(name[0..-2], arguments.first)
    else
      super
    end 
  end

  def get(name)
    recalculate if @need_to_recalculate
    raise NoMethodError.new(name) unless @outputs.has_key?(name)
    ruby_value_from_excel_value(@outputs[name])
  end

  def set(name, ruby_value)
    raise NoMethodError.new("#{name}=") unless INPUT_NAMES.has_key?(name)
    @need_to_recalculate = true
    @inputs[INPUT_NAMES[name]] = ruby_value
  end

  def recalculate
    c_inputs = @inputs.map { |r| excel_value_from_ruby_value(r) }
    pointer = C.run(*c_inputs)
    size = C::ExcelValue.size
    OUTPUT_NAMES.each.with_index do |ref, i|
      @outputs[ref] = C::ExcelValue.new(pointer + (i*size)) 
    end
    @need_to_recalculate = false
  end

  def ruby_value_from_excel_value(excel_value)
    case excel_value[:type]
    when :ExcelNumber; excel_value[:number]
    when :ExcelString; excel_value[:string].read_string.force_encoding("utf-8")
    when :ExcelBoolean; excel_value[:number] == 1
    when :ExcelEmpty; nil
    when :ExcelRange
      r = excel_value[:rows]
      c = excel_value[:columns]
      p = excel_value[:array]
      s = C::ExcelValue.size
      a = Array.new(r) { Array.new(c) }
      (0...r).each do |row|
        (0...c).each do |column|
          a[row][column] = ruby_value_from_excel_value(C::ExcelValue.new(p + (((row*c)+column)*s)))
        end
      end 
      return a
    when :ExcelError; [:value,:name,:div0,:ref,:na][excel_value[:number]]
    else
      raise Exception.new("ExcelValue type \u0023{excel_value[:type].inspect} not recognised")
    end
  end

  def excel_value_from_ruby_value(ruby_value, excel_value = C::ExcelValue.new)
    case ruby_value
    when Numeric
      excel_value[:type] = :ExcelNumber
      excel_value[:number] = ruby_value
    when String
      excel_value[:type] = :ExcelString
      excel_value[:string] = FFI::MemoryPointer.from_string(ruby_value.encode('utf-8'))
    when TrueClass, FalseClass
      excel_value[:type] = :ExcelBoolean
      excel_value[:number] = ruby_value ? 1 : 0
    when nil
      excel_value[:type] = :ExcelEmpty
    when Array
      excel_value[:type] = :ExcelRange
      # Presumed to be a row unless specified otherwise
      if ruby_value.first.is_a?(Array)
        excel_value[:rows] = ruby_value.size
        excel_value[:columns] = ruby_value.first.size
      else
        excel_value[:rows] = 1
        excel_value[:columns] = ruby_value.size
      end
      ruby_values = ruby_value.flatten
      pointer = FFI::MemoryPointer.new(C::ExcelValue, ruby_values.size)
      excel_value[:array] = pointer
      ruby_values.each.with_index do |v,i|
        excel_value_from_ruby_value(v, C::ExcelValue.new(pointer[i]))
      end
    when Symbol
      excel_value[:type] = :ExcelError
      excel_value[:number] = [:value, :name, :div0, :ref, :na].index(ruby_value)
    else
      raise Exception.new("Ruby value \u0023{ruby_value.inspect} not translatable into excel")
    end
    excel_value
  end


  module C 
    extend FFI::Library
    ffi_lib  File.join(File.dirname(__FILE__),FFI.map_library_name('#{name}'))
    ExcelType = enum :ExcelEmpty, :ExcelNumber, :ExcelString, :ExcelBoolean, :ExcelError, :ExcelRange
                  
    class ExcelValue < FFI::Struct
      layout :type, ExcelType,
             :number, :double,
             :string, :pointer,
             :array, :pointer,
             :rows, :int,
             :columns, :int             
    end
    
    # use this function to reset all cell values
    attach_function 'reset', [], :void
END
  
    o.puts code

    arguments =  @named_references_that_can_be_set_at_runtime.map { |ref| "ExcelValue.by_value" }.join(", ")

    o.puts "    attach_function 'run', [#{arguments}], :pointer"
    o.puts "  end # C module"  
    o.puts "end # #{ruby_module_name}"  
    close(o)
  end
  
  def write_tests
    log.info "Writing tests" 

    name = output_name.downcase
    o = output("test_#{name}.rb")    
    o.puts "# coding: utf-8"
    o.puts "# Test for #{name}"
    o.puts  "require 'minitest/autorun'"
    o.puts  "require_relative '#{output_name.downcase}'"
    o.puts
    o.puts "class Test#{ruby_module_name} < Minitest::Unit::TestCase"
    o.puts "  def test_run"
    o.puts "    workbook = #{ruby_module_name}.new"
    o.puts
    
    @named_references_that_can_be_set_at_runtime.each do |ref|
    o.puts "    workbook.#{c_name_for(ref)} = #{default_value_for_named_reference(ref).inspect}"
    end
    o.puts

    @named_references_to_keep.each do |ref|
    o.puts "    assert_equal(workbook.#{c_name_for(ref)}, #{default_value_for_named_reference(ref).inspect})"
    end
    o.puts
    
    o.puts "  end"

    o.puts "end"
    close(o)
  end

end
