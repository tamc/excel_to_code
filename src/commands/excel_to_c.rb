# coding: utf-8
require_relative 'excel_to_x'
require 'ffi'

class ExcelToC < ExcelToX
  
  def language
    "c"
  end
  
  # These actually create the code version of the excel
  def write_code
    write_out_excel_as_code
    write_build_script
    write_fuby_ffi_interface
    write_tests
  end
    
  def write_out_excel_as_code
        
    number_of_refs = @formulae.size
        
    # Output the workbook preamble
    o = output("#{output_name.downcase}.c")
    o.puts "// #{excel_file} approximately translated into C"

    o.puts '// First we have c versions of all the excel functions that we know'
    o.puts IO.readlines(File.join(File.dirname(__FILE__),'..','compile','c','excel_to_c_runtime.c')).join
    o.puts '// End of the generic c functions'
    o.puts 
    o.puts '// Start of the file specific functions'
    o.puts
    
    # Now we have to put all the initial definitions out
    o.puts "// definitions"
    o.puts "static ExcelValue ORIGINAL_EXCEL_FILENAME = {.type = ExcelString, .string = #{excel_file.inspect} };"

    c = CompileToCHeader.new
    c.settable = settable
    c.gettable = gettable
    c.rewrite(@formulae, @worksheet_c_names, o)
    
    # Need to make sure there are enough refs for named references as well
    number_of_refs += @named_references_to_keep.size

    o.puts "// end of definitions"
    o.puts
    o.puts "// Used to decide whether to recalculate a cell"
    o.puts "static int variable_set[#{number_of_refs}];"
    o.puts ""
    o.puts "// Used to reset all cached values and free up memory"
    # FIXME: This feels like a bad place for this. Should be in runtime?
    o.puts "void reset() {"
    o.puts "  int i;"
    o.puts "  cell_counter = 0;"
    o.puts "  free_all_allocated_memory(); "
    o.puts "  for(i = 0; i < #{number_of_refs}; i++) {"
    o.puts "    variable_set[i] = 0;"
    o.puts "  }"
    o.puts "};"
    o.puts
    
    # Output the value constants
    o.puts "// starting the value constants"
    mapper = MapValuesToCStructs.new
    @constants.each do |ref, ast|
      begin
        calculation = mapper.map(ast)
        o.puts "static ExcelValue #{ref} = #{calculation};"
      rescue Exception => e
        puts "Exception at #{ref} #{ast}"
        raise
      end
    end          
    o.puts "// ending the value constants"
    o.puts
    
    variable_set_counter = 0
    
    c = CompileToC.new
    c.variable_set_counter = variable_set_counter
    # Output the elements from each worksheet in turn
    c.settable = settable
    c.gettable = gettable
    c.rewrite(@formulae, @worksheet_c_names, o)

    # Output the named references

    # Getters
    o.puts "// Start of named references"
    c.gettable = lambda { |ref| true }
    c.settable = lambda { |ref| false }
    c.rewrite(@named_references_to_keep, @worksheet_c_names, o)

    # Setters
    # c = CompileNamedReferenceSetters.new
    # c.cells_that_can_be_set_at_runtime = cells_that_can_be_set_at_runtime
    # c.rewrite(i,w,o)

    # close(i)
    # o.puts "// End of named references"

    # close(w,o)
  end
  
  # FIXME: Should make a Rakefile, especially in order to make sure the dynamic library name
  # is set properly
  def write_build_script
    o = output("Makefile")
    name = output_name.downcase
    
    # Target for shared library
    shared_library_name = FFI.map_library_name(name)
    o.puts "#{shared_library_name}: #{name}.o"
    o.puts "\tgcc -shared -o #{shared_library_name} #{name}.o"
    o.puts
    
    # Target for compiled version
    o.puts "#{name}.o:"
    o.puts "\tgcc -Wall -fPIC -c #{name}.c"
    o.puts
    
    # Target for cleaning
    o.puts "clean:"
    o.puts "\trm #{name}.o"
    o.puts "\trm #{shared_library_name}"
    
    close(o)
  end
  
  def write_fuby_ffi_interface
    name = output_name.downcase
    o = output("#{name}.rb")
      
    code = <<END
require 'ffi'
require 'singleton'

class #{ruby_module_name}Shim

  # WARNING: this is not thread safe
  def initialize
    reset
  end

  def reset
    #{ruby_module_name}.reset
  end

  def method_missing(name, *arguments)
    if arguments.size == 0
      get(name)
    elsif arguments.size == 1
      set(name, arguments.first)
    else
      super
    end 
  end

  def get(name)
    return 0 unless #{ruby_module_name}.respond_to?(name)
    ruby_value_from_excel_value(#{ruby_module_name}.send(name))
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
      s = #{ruby_module_name}::ExcelValue.size
      a = Array.new(r) { Array.new(c) }
      (0...r).each do |row|
        (0...c).each do |column|
          a[row][column] = ruby_value_from_excel_value(#{ruby_module_name}::ExcelValue.new(p + (((row*c)+column)*s)))
        end
      end 
      return a
    when :ExcelError; [:value,:name,:div0,:ref,:na][excel_value[:number]]
    else
      raise Exception.new("ExcelValue type \u0023{excel_value[:type].inspect} not recognised")
    end
  end

  def set(name, ruby_value)
    name = name.to_s
    name = "set_\#{name[0..-2]}" if name.end_with?('=')
    return false unless #{ruby_module_name}.respond_to?(name)
    #{ruby_module_name}.send(name, excel_value_from_ruby_value(ruby_value))
  end

  def excel_value_from_ruby_value(ruby_value, excel_value = #{ruby_module_name}::ExcelValue.new)
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
      pointer = FFI::MemoryPointer.new(#{ruby_module_name}::ExcelValue, ruby_values.size)
      excel_value[:array] = pointer
      ruby_values.each.with_index do |v,i|
        excel_value_from_ruby_value(v, #{ruby_module_name}::ExcelValue.new(pointer[i]))
      end
    when Symbol
      excel_value[:type] = :ExcelError
      excel_value[:number] = [:value, :name, :div0, :ref, :na].index(ruby_value)
    else
      raise Exception.new("Ruby value \u0023{ruby_value.inspect} not translatable into excel")
    end
    excel_value
  end

end
    

module #{ruby_module_name}
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
  
END
    o.puts code
    o.puts
    o.puts "  # use this function to reset all cell values"
    o.puts "  attach_function 'reset', [], :void"


    worksheets do |name, xml_filename|
      c_name = c_name_for_worksheet_name(name)

      # Put in place the setters, if any
      settable_refs = @cells_that_can_be_set_at_runtime[name]
      if settable_refs
        settable_refs = all_formulae[name].keys if settable_refs == :all
        settable_refs.each do |ref|
          o.puts "  attach_function 'set_#{c_name}_#{ref.downcase}', [ExcelValue.by_value], :void"
        end
      end

      # Put in place the getters
      if !cells_to_keep || cells_to_keep.empty? || cells_to_keep[name] == :all
        getable_refs = @formulae.keys.select { |ref| ref.first == name }.map { |ref| ref.last } 
      elsif !cells_to_keep[name] && settable_refs
        getable_refs = settable_refs
      else
        getable_refs = cells_to_keep[name] || []
      end
              
      getable_refs.each do |ref|
        o.puts "  attach_function '#{c_name}_#{ref.downcase}', [], ExcelValue.by_value"
      end
        
      o.puts "  # end of #{name}"
    end

    # Now put in place the getters and setters for the named references
    # o.puts "  # Start of named references"

    # # Getters
    # i = input('Named references to keep')
    # i.each_line do |line|
    #   name = line.strip.split("\t").first
    #   o.puts "  attach_function '#{name}', [], ExcelValue.by_value"
    # end
    # close(i)

    # # Setters
    # i = input('Named references to set')
    # i.each_line do |line|
    #   name = line.strip.split("\t").first
    #   o.puts "  attach_function 'set_#{name}', [ExcelValue.by_value], :void"
    # end

    # close(i)
    # o.puts "  # End of named references"

    o.puts "end"  
    close(o)
  end
  
  def write_tests
    name = output_name.downcase
    o = output("test_#{name}.rb")    
    o.puts "# coding: utf-8"
    o.puts "# Test for #{name}"
    o.puts "require 'rubygems'"
    o.puts "gem 'minitest'"
    o.puts  "require 'minitest/autorun'"
    o.puts  "require_relative '#{output_name.downcase}'"
    o.puts
    o.puts "class Test#{ruby_module_name} < Minitest::Test"
    o.puts "  def self.runnable_methods"
    o.puts "    puts 'Overriding minitest to run tests in a defined order'"
    o.puts "    methods = methods_matching(/^test_/)"
    o.puts "  end" 
    o.puts "  def worksheet; @worksheet ||= init_spreadsheet; end"
    o.puts "  def init_spreadsheet; #{ruby_module_name}Shim.new end"
    
    CompileToCUnitTest.rewrite(Hash[@references_to_test_array], sloppy_tests, @worksheet_c_names, o)
    o.puts "end"
    close(o)
  end

  
  def compile_code
    return unless actually_compile_code || actually_run_tests
    puts "Compiling the resulting c code"
    puts `cd #{File.join(output_directory)}; make clean; make`
  end
  
  def run_tests
    return unless actually_run_tests
    puts "Running the resulting tests"
    puts `cd #{File.join(output_directory)}; ruby "test_#{output_name.downcase}.rb"`
  end
  
end
