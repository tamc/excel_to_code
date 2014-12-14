# coding: utf-8
require_relative 'excel_to_x'
require 'ffi'

class ExcelToC < ExcelToX

  # If true, creates a Rakefile, if false, doesn't (default true)
  attr_accessor :create_rakefile
  # If true, creates a Makefile, if false, doesn't (default false)
  attr_accessor :create_makefile
  # If true, writes tests in C rather than in ruby
  attr_accessor :write_tests_in_c
  
  def set_defaults
    super
    @create_rakefile = true if @create_rakefile == nil
    @create_makefile = false if @create_makefile == nil
    @write_tests_in_c = false if @write_tests_in_c == nil
  end

  def language
    "c"
  end
  
  # These actually create the code version of the excel
  def write_code
    write_out_excel_as_code
    write_build_script
    write_fuby_ffi_interface
    if write_tests_in_c
      write_tests_as_c
    else
      write_tests_as_ruby
    end
  end
    
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
    

    c = CompileToCHeader.new
    c.settable = settable
    c.gettable = gettable
    c.rewrite(@formulae, @worksheet_c_names, o)
    
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
    named_references_ast = {}
    @named_references_to_keep.each do |ref|
      c_name = ref.is_a?(Array) ? c_name_for(ref) : ["", c_name_for(ref)]
      named_references_ast[c_name] = @named_references[ref] || @table_areas[ref]
    end

    c.rewrite(named_references_ast, @worksheet_c_names, o)

    # Setters
    c = CompileNamedReferenceSetters.new
    c.cells_that_can_be_set_at_runtime = cells_that_can_be_set_at_runtime
    named_references_ast = {}
    @named_references_that_can_be_set_at_runtime.each do |ref|
      named_references_ast[c_name_for(ref)] = @named_references[ref] || @table_areas[ref]
    end
    c.rewrite(named_references_ast, @worksheet_c_names, o)
    o.puts "// End of named references"

    close(o)
  end
  
  # FIXME: Should make a Rakefile, especially in order to make sure the dynamic library name
  # is set properly
  def write_build_script
    log.info "Writing Build script"
    write_makefile if create_makefile
    write_rakefile if create_rakefile
  end

  def write_makefile
    log.info "Writing Makefile"

    name = output_name.downcase
    o = output("Makefile")

    # Target for shared library
    shared_library_name = FFI.map_library_name(name)
    o.puts "#{shared_library_name}: #{name}.o"
    o.puts "\tgcc -shared -o #{shared_library_name} #{name}.o"
    o.puts

    # Target for compiled version
    o.puts "#{name}.o:"
    o.puts "\tgcc -fPIC -c #{name}.c"
    o.puts

    # Target for cleaning
    o.puts "clean:"
    o.puts "\trm #{name}.o"
    o.puts "\trm #{shared_library_name}"

    close(o)
  end

  def write_rakefile
    log.info "Writing Rakefile"

    o = output("Rakefile")
    name = output_name.downcase
    o.puts "require 'ffi'"
    o.puts

    o.puts "this_directory = File.dirname(__FILE__)"
    o.puts
    o.puts "COMPILER = 'gcc'"
    o.puts "COMPILE_FLAGS = '-fPIC'"
    o.puts "SHARED_LIBRARY_FLAGS = '-shared -fPIC'"
    o.puts

    o.puts "OUTPUT = FFI.map_library_name '#{name}'"
    o.puts "OUTPUT_DIR = this_directory" 
    o.puts "SOURCE = '#{name}.c'"
    o.puts "OBJECT = '#{name}.o'"
    o.puts

    o.puts "task :default => [:build]"
    o.puts
    o.puts "desc 'Build the 2050 model, then install it'"
    o.puts "task :build => [OUTPUT]"
    o.puts

    o.puts "file OUTPUT => OBJECT do"
    o.puts '  puts "Turning #{OBJECT} and putting it in #{OUTPUT_DIR} as #{OUTPUT}"'
    o.puts '  puts "Note that this is a really large c file, it may take tens of minutes to compile."'
    o.puts '  sh "#{COMPILER} #{SHARED_LIBRARY_FLAGS} -o #{File.join(OUTPUT_DIR,OUTPUT)} #{OBJECT}"'
    o.puts 'end'
    o.puts

    o.puts 'file OBJECT => SOURCE do'
    o.puts '  puts "Building #{SOURCE}"'
    o.puts '  puts "Note that this is a really large c file, it may take tens of minutes to compile."'
    o.puts '  sh "#{COMPILER} #{COMPILE_FLAGS} -o #{OBJECT} -c #{SOURCE}"'
    o.puts 'end'
  end
  
  def write_fuby_ffi_interface
    log.info "Writing ruby FFI code"

    name = output_name.downcase
    o = output("#{name}.rb")
      
    code = <<END
require 'ffi'
require 'singleton'

class #{ruby_module_name}

  # WARNING: this is not thread safe
  def initialize
    reset
  end

  def reset
    C.reset
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
    return 0 unless C.respond_to?(name)
    ruby_value_from_excel_value(C.send(name))
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

  def set(name, ruby_value)
    name = name.to_s
    name = "set_\#{name[0..-2]}" if name.end_with?('=')
    return false unless C.respond_to?(name)
    C.send(name, excel_value_from_ruby_value(ruby_value))
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
  
END
    o.puts code
    o.puts
    o.puts "    # use this function to reset all cell values"
    o.puts "    attach_function 'reset', [], :void"


    worksheets do |name, xml_filename|
      c_name = c_name_for_worksheet_name(name)

      # Put in place the setters, if any
      settable_refs = @cells_that_can_be_set_at_runtime[name]
      if settable_refs
        settable_refs = @formulae.keys.select { |k| k.first == name }.map { |k| k.last } if settable_refs == :all
        settable_refs.each do |ref|
          o.puts "    attach_function 'set_#{c_name}_#{ref.downcase}', [ExcelValue.by_value], :void"
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
        o.puts "    attach_function '#{c_name}_#{ref.downcase}', [], ExcelValue.by_value"
      end
        
      o.puts "    # end of #{name}"
    end

    o.puts "    # Start of named references"
    # Getters
    @named_references_to_keep.each do |name|
      o.puts "    attach_function '#{c_name_for(name)}', [], ExcelValue.by_value"
    end

    # Setters
    @named_references_that_can_be_set_at_runtime.each do |name|
      o.puts "    attach_function 'set_#{c_name_for(name)}', [ExcelValue.by_value], :void"
    end

    o.puts "    # End of named references"

    o.puts "  end # C module"  
    o.puts "end # #{ruby_module_name}"  
    close(o)
  end
  
  def write_tests_as_ruby
    log.info "Writing tests in ruby" 

    name = output_name.downcase
    o = output("test_#{name}.rb")    
    o.puts "# coding: utf-8"
    o.puts "# Test for #{name}"
    o.puts  "require 'minitest/autorun'"
    o.puts  "require_relative '#{output_name.downcase}'"
    o.puts
    o.puts "class Test#{ruby_module_name} < Minitest::Unit::TestCase"
    o.puts "  def self.runnable_methods"
    o.puts "    puts 'Overriding minitest to run tests in a defined order'"
    o.puts "    methods = methods_matching(/^test_/)"
    o.puts "  end" 
    o.puts "  def worksheet; @worksheet ||= init_spreadsheet; end"
    o.puts "  def init_spreadsheet; #{ruby_module_name}.new end"
    
    CompileToRubyUnitTest.rewrite(Hash[@references_to_test_array], sloppy_tests, @worksheet_c_names, @constants, o)
    o.puts "end"
    close(o)
  end

  def write_tests_as_c
    log.info "Writing tests in C" 

    name = output_name.downcase
    o = output("test_#{name}.c")    
    o.puts "#include \"#{name}.c\""
    o.puts "int main() {"
    o.puts "  printf(\"\\n\\nRunning tests on #{name}\\n\\n\");"
    CompileToCUnitTest.rewrite(Hash[@references_to_test_array], sloppy_tests, @worksheet_c_names, @constants, o)
    o.puts "  printf(\"\\n\\nFinished tests on #{name}\\n\\n\");"
    o.puts "  return 0;"
    o.puts "}"
    close(o)
  end
  
  
  def compile_code
    return unless actually_compile_code || actually_run_tests
    name = output_name.downcase
    log.info "Compiling"
    puts `gcc -fPIC -o #{File.join(output_directory, name)}.o -c #{File.join(output_directory, name)}.c`
    puts `gcc -shared -fPIC -o #{File.join(output_directory, FFI.map_library_name(name))} #{File.join(output_directory, name)}.o`
  end
  
  def run_tests
    return unless actually_run_tests
    puts "Running the resulting tests"
    if write_tests_as_c
      puts `cd #{File.join(output_directory)}; gcc "test_#{output_name.downcase}.c"; ./a.out`
    else
      puts `cd #{File.join(output_directory)}; ruby "test_#{output_name.downcase}.rb"`
    end
  end
  
end
