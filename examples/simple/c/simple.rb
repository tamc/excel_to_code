require 'ffi'
require 'singleton'

class Simple

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
    when :ExcelError; [:value,:name,:div0,:ref,:na,:num][excel_value[:number]]
    else
      raise Exception.new("ExcelValue type #{excel_value[:type].inspect} not recognised")
    end
  end

  def set(name, ruby_value)
    name = name.to_s
    name = "set_#{name[0..-2]}" if name.end_with?('=')
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
      raise Exception.new("Ruby value #{ruby_value.inspect} not translatable into excel")
    end
    excel_value
  end


  module C 
    extend FFI::Library
    ffi_lib  File.join(File.dirname(__FILE__),FFI.map_library_name('simple'))
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
    attach_function 'set_inputs_b1', [ExcelValue.by_value], :void
    attach_function 'set_inputs_b2', [ExcelValue.by_value], :void
    attach_function 'set_inputs_b4', [ExcelValue.by_value], :void
    attach_function 'inputs_a1', [], ExcelValue.by_value
    attach_function 'inputs_b1', [], ExcelValue.by_value
    attach_function 'inputs_a2', [], ExcelValue.by_value
    attach_function 'inputs_b2', [], ExcelValue.by_value
    attach_function 'inputs_a4', [], ExcelValue.by_value
    attach_function 'inputs_b4', [], ExcelValue.by_value
    # end of Inputs
    attach_function 'outputs_a1', [], ExcelValue.by_value
    attach_function 'outputs_a2', [], ExcelValue.by_value
    attach_function 'outputs_a3', [], ExcelValue.by_value
    # end of Outputs
    # Start of named references
    # End of named references
  end # C module
end # Simple
