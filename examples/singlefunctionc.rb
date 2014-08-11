require 'ffi'
require 'singleton'

class Singlefunctionc

  INPUT_NAMES = {"input"=>0}
  OUTPUT_NAMES = ["output", "input"]

  def initialize
    reset
  end

  def reset
    @need_to_recalculate = true
    @inputs = [100.0] # Defaults
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
    raise NoMethodError.new("singlefunctionc=") unless INPUT_NAMES.has_key?(name)
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
      raise Exception.new("ExcelValue type #{excel_value[:type].inspect} not recognised")
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
      raise Exception.new("Ruby value #{ruby_value.inspect} not translatable into excel")
    end
    excel_value
  end


  module C 
    extend FFI::Library
    ffi_lib  File.join(File.dirname(__FILE__),FFI.map_library_name('singlefunctionc'))
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
    attach_function 'run', [ExcelValue.by_value], :pointer
  end # C module
end # Singlefunctionc
