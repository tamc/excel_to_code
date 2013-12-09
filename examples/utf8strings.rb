require 'ffi'
require 'singleton'

class Utf8stringsShim

  # WARNING: this is not thread safe
  def initialize
    reset
  end

  def reset
    Utf8strings.reset
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
    return 0 unless Utf8strings.respond_to?(name)
    ruby_value_from_excel_value(Utf8strings.send(name))
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
      s = Utf8strings::ExcelValue.size
      a = Array.new(r) { Array.new(c) }
      (0...r).each do |row|
        (0...c).each do |column|
          a[row][column] = ruby_value_from_excel_value(Utf8strings::ExcelValue.new(p + (((row*c)+column)*s)))
        end
      end 
      return a
    when :ExcelError; [:value,:name,:div0,:ref,:na][excel_value[:number]]
    else
      raise Exception.new("ExcelValue type #{excel_value[:type].inspect} not recognised")
    end
  end

  def set(name, ruby_value)
    name = name.to_s
    name = "set_#{name[0..-2]}" if name.end_with?('=')
    return false unless Utf8strings.respond_to?(name)
    Utf8strings.send(name, excel_value_from_ruby_value(ruby_value))
  end

  def excel_value_from_ruby_value(ruby_value, excel_value = Utf8strings::ExcelValue.new)
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
      pointer = FFI::MemoryPointer.new(Utf8strings::ExcelValue, ruby_values.size)
      excel_value[:array] = pointer
      ruby_values.each.with_index do |v,i|
        excel_value_from_ruby_value(v, Utf8strings::ExcelValue.new(pointer[i]))
      end
    when Symbol
      excel_value[:type] = :ExcelError
      excel_value[:number] = [:value, :name, :div0, :ref, :na].index(ruby_value)
    else
      raise Exception.new("Ruby value #{ruby_value.inspect} not translatable into excel")
    end
    excel_value
  end

end
    

module Utf8strings
  extend FFI::Library
  ffi_lib  File.join(File.dirname(__FILE__),FFI.map_library_name('utf8strings'))
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
  attach_function 'sheet1_a1', [], ExcelValue.by_value
  attach_function 'sheet1_b1', [], ExcelValue.by_value
  attach_function 'sheet1_a2', [], ExcelValue.by_value
  attach_function 'sheet1_b2', [], ExcelValue.by_value
  attach_function 'sheet1_a3', [], ExcelValue.by_value
  attach_function 'sheet1_b3', [], ExcelValue.by_value
  # end of Sheet1
  attach_function 's_1_a2', [], ExcelValue.by_value
  attach_function 's_1_c2', [], ExcelValue.by_value
  attach_function 's_1_b4', [], ExcelValue.by_value
  attach_function 's_1_b5', [], ExcelValue.by_value
  attach_function 's_1_b6', [], ExcelValue.by_value
  attach_function 's_1_b9', [], ExcelValue.by_value
  attach_function 's_1_b11', [], ExcelValue.by_value
  # end of 工作表1
  attach_function 'set_s2015_b3', [ExcelValue.by_value], :void
  attach_function 's2015_b3', [], ExcelValue.by_value
  # end of 2015
  # Start of named references
  # End of named references
end
