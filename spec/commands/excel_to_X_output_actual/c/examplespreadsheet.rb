require 'ffi'
require 'singleton'

class ExampleSpreadsheetShim

  # WARNING: this is not thread safe
  def initialize
    reset
  end

  def reset
    ExampleSpreadsheet.reset
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
    return 0 unless ExampleSpreadsheet.respond_to?(name)
    ruby_value_from_excel_value(ExampleSpreadsheet.send(name))
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
      s = ExampleSpreadsheet::ExcelValue.size
      a = Array.new(r) { Array.new(c) }
      (0...r).each do |row|
        (0...c).each do |column|
          a[row][column] = ruby_value_from_excel_value(ExampleSpreadsheet::ExcelValue.new(p + (((row*c)+column)*s)))
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
    return false unless ExampleSpreadsheet.respond_to?(name)
    ExampleSpreadsheet.send(name, excel_value_from_ruby_value(ruby_value))
  end

  def excel_value_from_ruby_value(ruby_value, excel_value = ExampleSpreadsheet::ExcelValue.new)
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
      pointer = FFI::MemoryPointer.new(ExampleSpreadsheet::ExcelValue, ruby_values.size)
      excel_value[:array] = pointer
      ruby_values.each.with_index do |v,i|
        excel_value_from_ruby_value(v, ExampleSpreadsheet::ExcelValue.new(pointer[i]))
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
    

module ExampleSpreadsheet
  extend FFI::Library
  ffi_lib  File.join(File.dirname(__FILE__),FFI.map_library_name('examplespreadsheet'))
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
  attach_function 'set_valuetypes_a1', [ExcelValue.by_value], :void
  attach_function 'set_valuetypes_a2', [ExcelValue.by_value], :void
  attach_function 'set_valuetypes_a3', [ExcelValue.by_value], :void
  attach_function 'set_valuetypes_a4', [ExcelValue.by_value], :void
  attach_function 'set_valuetypes_a5', [ExcelValue.by_value], :void
  attach_function 'set_valuetypes_a6', [ExcelValue.by_value], :void
  attach_function 'valuetypes_a1', [], ExcelValue.by_value
  attach_function 'valuetypes_a2', [], ExcelValue.by_value
  attach_function 'valuetypes_a3', [], ExcelValue.by_value
  attach_function 'valuetypes_a4', [], ExcelValue.by_value
  attach_function 'valuetypes_a5', [], ExcelValue.by_value
  attach_function 'valuetypes_a6', [], ExcelValue.by_value
  # end of ValueTypes
  attach_function 'formulaetypes_a1', [], ExcelValue.by_value
  attach_function 'formulaetypes_b1', [], ExcelValue.by_value
  attach_function 'formulaetypes_a2', [], ExcelValue.by_value
  attach_function 'formulaetypes_b2', [], ExcelValue.by_value
  attach_function 'formulaetypes_a3', [], ExcelValue.by_value
  attach_function 'formulaetypes_b3', [], ExcelValue.by_value
  attach_function 'formulaetypes_a4', [], ExcelValue.by_value
  attach_function 'formulaetypes_b4', [], ExcelValue.by_value
  attach_function 'formulaetypes_a5', [], ExcelValue.by_value
  attach_function 'formulaetypes_b5', [], ExcelValue.by_value
  attach_function 'formulaetypes_a6', [], ExcelValue.by_value
  attach_function 'formulaetypes_b6', [], ExcelValue.by_value
  attach_function 'formulaetypes_a7', [], ExcelValue.by_value
  attach_function 'formulaetypes_b7', [], ExcelValue.by_value
  attach_function 'formulaetypes_a8', [], ExcelValue.by_value
  attach_function 'formulaetypes_b8', [], ExcelValue.by_value
  # end of FormulaeTypes
  attach_function 'set_ranges_f4', [ExcelValue.by_value], :void
  attach_function 'set_ranges_f5', [ExcelValue.by_value], :void
  attach_function 'set_ranges_f6', [ExcelValue.by_value], :void
  attach_function 'set_ranges_e5', [ExcelValue.by_value], :void
  attach_function 'set_ranges_g5', [ExcelValue.by_value], :void
  attach_function 'ranges_b1', [], ExcelValue.by_value
  attach_function 'ranges_c1', [], ExcelValue.by_value
  attach_function 'ranges_a2', [], ExcelValue.by_value
  attach_function 'ranges_b2', [], ExcelValue.by_value
  attach_function 'ranges_c2', [], ExcelValue.by_value
  attach_function 'ranges_a3', [], ExcelValue.by_value
  attach_function 'ranges_b3', [], ExcelValue.by_value
  attach_function 'ranges_c3', [], ExcelValue.by_value
  attach_function 'ranges_a4', [], ExcelValue.by_value
  attach_function 'ranges_b4', [], ExcelValue.by_value
  attach_function 'ranges_c4', [], ExcelValue.by_value
  attach_function 'ranges_f4', [], ExcelValue.by_value
  attach_function 'ranges_e5', [], ExcelValue.by_value
  attach_function 'ranges_f5', [], ExcelValue.by_value
  attach_function 'ranges_g5', [], ExcelValue.by_value
  attach_function 'ranges_f6', [], ExcelValue.by_value
  # end of Ranges
  attach_function 'set_referencing_a4', [ExcelValue.by_value], :void
  attach_function 'set_referencing_c15', [ExcelValue.by_value], :void
  attach_function 'set_referencing_d15', [ExcelValue.by_value], :void
  attach_function 'set_referencing_e15', [ExcelValue.by_value], :void
  attach_function 'set_referencing_f15', [ExcelValue.by_value], :void
  attach_function 'set_referencing_c16', [ExcelValue.by_value], :void
  attach_function 'set_referencing_d16', [ExcelValue.by_value], :void
  attach_function 'set_referencing_e16', [ExcelValue.by_value], :void
  attach_function 'set_referencing_f16', [ExcelValue.by_value], :void
  attach_function 'set_referencing_c17', [ExcelValue.by_value], :void
  attach_function 'set_referencing_d17', [ExcelValue.by_value], :void
  attach_function 'set_referencing_e17', [ExcelValue.by_value], :void
  attach_function 'set_referencing_f17', [ExcelValue.by_value], :void
  attach_function 'set_referencing_c18', [ExcelValue.by_value], :void
  attach_function 'set_referencing_d18', [ExcelValue.by_value], :void
  attach_function 'set_referencing_e18', [ExcelValue.by_value], :void
  attach_function 'set_referencing_f18', [ExcelValue.by_value], :void
  attach_function 'set_referencing_c19', [ExcelValue.by_value], :void
  attach_function 'set_referencing_d19', [ExcelValue.by_value], :void
  attach_function 'set_referencing_e19', [ExcelValue.by_value], :void
  attach_function 'set_referencing_f19', [ExcelValue.by_value], :void
  attach_function 'set_referencing_c22', [ExcelValue.by_value], :void
  attach_function 'set_referencing_f34', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g34', [ExcelValue.by_value], :void
  attach_function 'set_referencing_h34', [ExcelValue.by_value], :void
  attach_function 'set_referencing_i34', [ExcelValue.by_value], :void
  attach_function 'set_referencing_j34', [ExcelValue.by_value], :void
  attach_function 'set_referencing_k34', [ExcelValue.by_value], :void
  attach_function 'set_referencing_l34', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m34', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n34', [ExcelValue.by_value], :void
  attach_function 'set_referencing_j35', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m35', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n35', [ExcelValue.by_value], :void
  attach_function 'set_referencing_j36', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m36', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n36', [ExcelValue.by_value], :void
  attach_function 'set_referencing_f37', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m37', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n37', [ExcelValue.by_value], :void
  attach_function 'set_referencing_i38', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m38', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n38', [ExcelValue.by_value], :void
  attach_function 'set_referencing_h39', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m39', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n39', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g40', [ExcelValue.by_value], :void
  attach_function 'set_referencing_j40', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m40', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n40', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g41', [ExcelValue.by_value], :void
  attach_function 'set_referencing_j41', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m41', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n41', [ExcelValue.by_value], :void
  attach_function 'set_referencing_f42', [ExcelValue.by_value], :void
  attach_function 'set_referencing_l42', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m42', [ExcelValue.by_value], :void
  attach_function 'set_referencing_f43', [ExcelValue.by_value], :void
  attach_function 'set_referencing_l43', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m43', [ExcelValue.by_value], :void
  attach_function 'set_referencing_l44', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m44', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n44', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g45', [ExcelValue.by_value], :void
  attach_function 'set_referencing_j45', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m45', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n45', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g46', [ExcelValue.by_value], :void
  attach_function 'set_referencing_h46', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m46', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n46', [ExcelValue.by_value], :void
  attach_function 'set_referencing_k47', [ExcelValue.by_value], :void
  attach_function 'set_referencing_m47', [ExcelValue.by_value], :void
  attach_function 'set_referencing_n47', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g50', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g51', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g52', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g53', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g54', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g55', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g56', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g57', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g58', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g59', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g60', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g61', [ExcelValue.by_value], :void
  attach_function 'set_referencing_g62', [ExcelValue.by_value], :void
  attach_function 'set_referencing_d64', [ExcelValue.by_value], :void
  attach_function 'referencing_a1', [], ExcelValue.by_value
  attach_function 'referencing_a2', [], ExcelValue.by_value
  attach_function 'referencing_a4', [], ExcelValue.by_value
  attach_function 'referencing_b4', [], ExcelValue.by_value
  attach_function 'referencing_c4', [], ExcelValue.by_value
  attach_function 'referencing_a5', [], ExcelValue.by_value
  attach_function 'referencing_b8', [], ExcelValue.by_value
  attach_function 'referencing_b9', [], ExcelValue.by_value
  attach_function 'referencing_b11', [], ExcelValue.by_value
  attach_function 'referencing_c11', [], ExcelValue.by_value
  attach_function 'referencing_c15', [], ExcelValue.by_value
  attach_function 'referencing_d15', [], ExcelValue.by_value
  attach_function 'referencing_e15', [], ExcelValue.by_value
  attach_function 'referencing_f15', [], ExcelValue.by_value
  attach_function 'referencing_c16', [], ExcelValue.by_value
  attach_function 'referencing_d16', [], ExcelValue.by_value
  attach_function 'referencing_e16', [], ExcelValue.by_value
  attach_function 'referencing_f16', [], ExcelValue.by_value
  attach_function 'referencing_c17', [], ExcelValue.by_value
  attach_function 'referencing_d17', [], ExcelValue.by_value
  attach_function 'referencing_e17', [], ExcelValue.by_value
  attach_function 'referencing_f17', [], ExcelValue.by_value
  attach_function 'referencing_c18', [], ExcelValue.by_value
  attach_function 'referencing_d18', [], ExcelValue.by_value
  attach_function 'referencing_e18', [], ExcelValue.by_value
  attach_function 'referencing_f18', [], ExcelValue.by_value
  attach_function 'referencing_c19', [], ExcelValue.by_value
  attach_function 'referencing_d19', [], ExcelValue.by_value
  attach_function 'referencing_e19', [], ExcelValue.by_value
  attach_function 'referencing_f19', [], ExcelValue.by_value
  attach_function 'referencing_c22', [], ExcelValue.by_value
  attach_function 'referencing_d22', [], ExcelValue.by_value
  attach_function 'referencing_d23', [], ExcelValue.by_value
  attach_function 'referencing_d24', [], ExcelValue.by_value
  attach_function 'referencing_d25', [], ExcelValue.by_value
  attach_function 'referencing_c31', [], ExcelValue.by_value
  attach_function 'referencing_o31', [], ExcelValue.by_value
  attach_function 'referencing_f33', [], ExcelValue.by_value
  attach_function 'referencing_g33', [], ExcelValue.by_value
  attach_function 'referencing_h33', [], ExcelValue.by_value
  attach_function 'referencing_i33', [], ExcelValue.by_value
  attach_function 'referencing_j33', [], ExcelValue.by_value
  attach_function 'referencing_k33', [], ExcelValue.by_value
  attach_function 'referencing_l33', [], ExcelValue.by_value
  attach_function 'referencing_m33', [], ExcelValue.by_value
  attach_function 'referencing_n33', [], ExcelValue.by_value
  attach_function 'referencing_o33', [], ExcelValue.by_value
  attach_function 'referencing_c34', [], ExcelValue.by_value
  attach_function 'referencing_d34', [], ExcelValue.by_value
  attach_function 'referencing_e34', [], ExcelValue.by_value
  attach_function 'referencing_f34', [], ExcelValue.by_value
  attach_function 'referencing_g34', [], ExcelValue.by_value
  attach_function 'referencing_h34', [], ExcelValue.by_value
  attach_function 'referencing_i34', [], ExcelValue.by_value
  attach_function 'referencing_j34', [], ExcelValue.by_value
  attach_function 'referencing_k34', [], ExcelValue.by_value
  attach_function 'referencing_l34', [], ExcelValue.by_value
  attach_function 'referencing_m34', [], ExcelValue.by_value
  attach_function 'referencing_n34', [], ExcelValue.by_value
  attach_function 'referencing_c35', [], ExcelValue.by_value
  attach_function 'referencing_d35', [], ExcelValue.by_value
  attach_function 'referencing_j35', [], ExcelValue.by_value
  attach_function 'referencing_m35', [], ExcelValue.by_value
  attach_function 'referencing_n35', [], ExcelValue.by_value
  attach_function 'referencing_o35', [], ExcelValue.by_value
  attach_function 'referencing_c36', [], ExcelValue.by_value
  attach_function 'referencing_d36', [], ExcelValue.by_value
  attach_function 'referencing_j36', [], ExcelValue.by_value
  attach_function 'referencing_m36', [], ExcelValue.by_value
  attach_function 'referencing_n36', [], ExcelValue.by_value
  attach_function 'referencing_o36', [], ExcelValue.by_value
  attach_function 'referencing_c37', [], ExcelValue.by_value
  attach_function 'referencing_d37', [], ExcelValue.by_value
  attach_function 'referencing_f37', [], ExcelValue.by_value
  attach_function 'referencing_m37', [], ExcelValue.by_value
  attach_function 'referencing_n37', [], ExcelValue.by_value
  attach_function 'referencing_o37', [], ExcelValue.by_value
  attach_function 'referencing_c38', [], ExcelValue.by_value
  attach_function 'referencing_d38', [], ExcelValue.by_value
  attach_function 'referencing_i38', [], ExcelValue.by_value
  attach_function 'referencing_m38', [], ExcelValue.by_value
  attach_function 'referencing_n38', [], ExcelValue.by_value
  attach_function 'referencing_o38', [], ExcelValue.by_value
  attach_function 'referencing_c39', [], ExcelValue.by_value
  attach_function 'referencing_d39', [], ExcelValue.by_value
  attach_function 'referencing_e39', [], ExcelValue.by_value
  attach_function 'referencing_h39', [], ExcelValue.by_value
  attach_function 'referencing_m39', [], ExcelValue.by_value
  attach_function 'referencing_n39', [], ExcelValue.by_value
  attach_function 'referencing_o39', [], ExcelValue.by_value
  attach_function 'referencing_c40', [], ExcelValue.by_value
  attach_function 'referencing_d40', [], ExcelValue.by_value
  attach_function 'referencing_e40', [], ExcelValue.by_value
  attach_function 'referencing_g40', [], ExcelValue.by_value
  attach_function 'referencing_j40', [], ExcelValue.by_value
  attach_function 'referencing_m40', [], ExcelValue.by_value
  attach_function 'referencing_n40', [], ExcelValue.by_value
  attach_function 'referencing_o40', [], ExcelValue.by_value
  attach_function 'referencing_c41', [], ExcelValue.by_value
  attach_function 'referencing_d41', [], ExcelValue.by_value
  attach_function 'referencing_e41', [], ExcelValue.by_value
  attach_function 'referencing_g41', [], ExcelValue.by_value
  attach_function 'referencing_j41', [], ExcelValue.by_value
  attach_function 'referencing_m41', [], ExcelValue.by_value
  attach_function 'referencing_n41', [], ExcelValue.by_value
  attach_function 'referencing_o41', [], ExcelValue.by_value
  attach_function 'referencing_c42', [], ExcelValue.by_value
  attach_function 'referencing_d42', [], ExcelValue.by_value
  attach_function 'referencing_f42', [], ExcelValue.by_value
  attach_function 'referencing_l42', [], ExcelValue.by_value
  attach_function 'referencing_m42', [], ExcelValue.by_value
  attach_function 'referencing_o42', [], ExcelValue.by_value
  attach_function 'referencing_c43', [], ExcelValue.by_value
  attach_function 'referencing_d43', [], ExcelValue.by_value
  attach_function 'referencing_f43', [], ExcelValue.by_value
  attach_function 'referencing_l43', [], ExcelValue.by_value
  attach_function 'referencing_m43', [], ExcelValue.by_value
  attach_function 'referencing_o43', [], ExcelValue.by_value
  attach_function 'referencing_c44', [], ExcelValue.by_value
  attach_function 'referencing_d44', [], ExcelValue.by_value
  attach_function 'referencing_l44', [], ExcelValue.by_value
  attach_function 'referencing_m44', [], ExcelValue.by_value
  attach_function 'referencing_n44', [], ExcelValue.by_value
  attach_function 'referencing_o44', [], ExcelValue.by_value
  attach_function 'referencing_c45', [], ExcelValue.by_value
  attach_function 'referencing_d45', [], ExcelValue.by_value
  attach_function 'referencing_g45', [], ExcelValue.by_value
  attach_function 'referencing_j45', [], ExcelValue.by_value
  attach_function 'referencing_m45', [], ExcelValue.by_value
  attach_function 'referencing_n45', [], ExcelValue.by_value
  attach_function 'referencing_o45', [], ExcelValue.by_value
  attach_function 'referencing_c46', [], ExcelValue.by_value
  attach_function 'referencing_d46', [], ExcelValue.by_value
  attach_function 'referencing_g46', [], ExcelValue.by_value
  attach_function 'referencing_h46', [], ExcelValue.by_value
  attach_function 'referencing_m46', [], ExcelValue.by_value
  attach_function 'referencing_n46', [], ExcelValue.by_value
  attach_function 'referencing_o46', [], ExcelValue.by_value
  attach_function 'referencing_c47', [], ExcelValue.by_value
  attach_function 'referencing_d47', [], ExcelValue.by_value
  attach_function 'referencing_e47', [], ExcelValue.by_value
  attach_function 'referencing_k47', [], ExcelValue.by_value
  attach_function 'referencing_m47', [], ExcelValue.by_value
  attach_function 'referencing_n47', [], ExcelValue.by_value
  attach_function 'referencing_o47', [], ExcelValue.by_value
  attach_function 'referencing_d50', [], ExcelValue.by_value
  attach_function 'referencing_g50', [], ExcelValue.by_value
  attach_function 'referencing_d51', [], ExcelValue.by_value
  attach_function 'referencing_g51', [], ExcelValue.by_value
  attach_function 'referencing_d52', [], ExcelValue.by_value
  attach_function 'referencing_g52', [], ExcelValue.by_value
  attach_function 'referencing_d53', [], ExcelValue.by_value
  attach_function 'referencing_g53', [], ExcelValue.by_value
  attach_function 'referencing_d54', [], ExcelValue.by_value
  attach_function 'referencing_g54', [], ExcelValue.by_value
  attach_function 'referencing_d55', [], ExcelValue.by_value
  attach_function 'referencing_g55', [], ExcelValue.by_value
  attach_function 'referencing_d56', [], ExcelValue.by_value
  attach_function 'referencing_g56', [], ExcelValue.by_value
  attach_function 'referencing_d57', [], ExcelValue.by_value
  attach_function 'referencing_g57', [], ExcelValue.by_value
  attach_function 'referencing_d58', [], ExcelValue.by_value
  attach_function 'referencing_g58', [], ExcelValue.by_value
  attach_function 'referencing_d59', [], ExcelValue.by_value
  attach_function 'referencing_g59', [], ExcelValue.by_value
  attach_function 'referencing_d60', [], ExcelValue.by_value
  attach_function 'referencing_g60', [], ExcelValue.by_value
  attach_function 'referencing_d61', [], ExcelValue.by_value
  attach_function 'referencing_g61', [], ExcelValue.by_value
  attach_function 'referencing_d62', [], ExcelValue.by_value
  attach_function 'referencing_g62', [], ExcelValue.by_value
  attach_function 'referencing_d64', [], ExcelValue.by_value
  attach_function 'referencing_e64', [], ExcelValue.by_value
  attach_function 'referencing_h64', [], ExcelValue.by_value
  # end of Referencing
  attach_function 'set_tables_b2', [ExcelValue.by_value], :void
  attach_function 'set_tables_c2', [ExcelValue.by_value], :void
  attach_function 'set_tables_d2', [ExcelValue.by_value], :void
  attach_function 'set_tables_b3', [ExcelValue.by_value], :void
  attach_function 'set_tables_c3', [ExcelValue.by_value], :void
  attach_function 'set_tables_b4', [ExcelValue.by_value], :void
  attach_function 'set_tables_c4', [ExcelValue.by_value], :void
  attach_function 'tables_a1', [], ExcelValue.by_value
  attach_function 'tables_b2', [], ExcelValue.by_value
  attach_function 'tables_c2', [], ExcelValue.by_value
  attach_function 'tables_d2', [], ExcelValue.by_value
  attach_function 'tables_b3', [], ExcelValue.by_value
  attach_function 'tables_c3', [], ExcelValue.by_value
  attach_function 'tables_d3', [], ExcelValue.by_value
  attach_function 'tables_b4', [], ExcelValue.by_value
  attach_function 'tables_c4', [], ExcelValue.by_value
  attach_function 'tables_d4', [], ExcelValue.by_value
  attach_function 'tables_f4', [], ExcelValue.by_value
  attach_function 'tables_g4', [], ExcelValue.by_value
  attach_function 'tables_h4', [], ExcelValue.by_value
  attach_function 'tables_b5', [], ExcelValue.by_value
  attach_function 'tables_c5', [], ExcelValue.by_value
  attach_function 'tables_e6', [], ExcelValue.by_value
  attach_function 'tables_f6', [], ExcelValue.by_value
  attach_function 'tables_g6', [], ExcelValue.by_value
  attach_function 'tables_e7', [], ExcelValue.by_value
  attach_function 'tables_f7', [], ExcelValue.by_value
  attach_function 'tables_g7', [], ExcelValue.by_value
  attach_function 'tables_e8', [], ExcelValue.by_value
  attach_function 'tables_f8', [], ExcelValue.by_value
  attach_function 'tables_g8', [], ExcelValue.by_value
  attach_function 'tables_e9', [], ExcelValue.by_value
  attach_function 'tables_f9', [], ExcelValue.by_value
  attach_function 'tables_g9', [], ExcelValue.by_value
  attach_function 'tables_c10', [], ExcelValue.by_value
  attach_function 'tables_e10', [], ExcelValue.by_value
  attach_function 'tables_f10', [], ExcelValue.by_value
  attach_function 'tables_g10', [], ExcelValue.by_value
  attach_function 'tables_c11', [], ExcelValue.by_value
  attach_function 'tables_e11', [], ExcelValue.by_value
  attach_function 'tables_f11', [], ExcelValue.by_value
  attach_function 'tables_g11', [], ExcelValue.by_value
  attach_function 'tables_c12', [], ExcelValue.by_value
  attach_function 'tables_c13', [], ExcelValue.by_value
  attach_function 'tables_c14', [], ExcelValue.by_value
  # end of Tables
  attach_function 's_innapropriate_sheet_name__c4', [], ExcelValue.by_value
  # end of (innapropriate) sheet name!
  # Start of named references
  # End of named references
end
