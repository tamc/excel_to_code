require 'ffi'
require 'singleton'

class OffsetsShim

  # WARNING: this is not thread safe
  def initialize
    reset
  end

  def reset
    Offsets.reset
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
    return 0 unless Offsets.respond_to?(name)
    ruby_value_from_excel_value(Offsets.send(name))
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
      s = Offsets::ExcelValue.size
      a = Array.new(r) { Array.new(c) }
      (0...r).each do |row|
        (0...c).each do |column|
          a[row][column] = ruby_value_from_excel_value(Offsets::ExcelValue.new(p + (((row*c)+column)*s)))
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
    return false unless Offsets.respond_to?(name)
    Offsets.send(name, excel_value_from_ruby_value(ruby_value))
  end

  def excel_value_from_ruby_value(ruby_value, excel_value = Offsets::ExcelValue.new)
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
      pointer = FFI::MemoryPointer.new(Offsets::ExcelValue, ruby_values.size)
      excel_value[:array] = pointer
      ruby_values.each.with_index do |v,i|
        excel_value_from_ruby_value(v, Offsets::ExcelValue.new(pointer[i]))
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
    

module Offsets
  extend FFI::Library
  ffi_lib  File.join(File.dirname(__FILE__),FFI.map_library_name('offsets'))
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

  # start of Model
  attach_function 'set_model_c22', [ExcelValue.by_value], :void
  attach_function 'set_model_c46', [ExcelValue.by_value], :void
  attach_function 'set_model_c24', [ExcelValue.by_value], :void
  attach_function 'set_model_c47', [ExcelValue.by_value], :void
  attach_function 'set_model_c26', [ExcelValue.by_value], :void
  attach_function 'set_model_c25', [ExcelValue.by_value], :void
  attach_function 'set_model_c48', [ExcelValue.by_value], :void
  attach_function 'set_model_c43', [ExcelValue.by_value], :void
  attach_function 'set_model_c42', [ExcelValue.by_value], :void
  attach_function 'set_model_c41', [ExcelValue.by_value], :void
  attach_function 'set_model_c29', [ExcelValue.by_value], :void
  attach_function 'set_model_c34', [ExcelValue.by_value], :void
  attach_function 'set_model_c39', [ExcelValue.by_value], :void
  attach_function 'set_model_c38', [ExcelValue.by_value], :void
  attach_function 'set_model_c37', [ExcelValue.by_value], :void
  attach_function 'set_model_c30', [ExcelValue.by_value], :void
  attach_function 'set_model_c35', [ExcelValue.by_value], :void
  attach_function 'set_model_c23', [ExcelValue.by_value], :void
  attach_function 'set_model_c33', [ExcelValue.by_value], :void
  attach_function 'model_c22', [], ExcelValue.by_value
  attach_function 'model_c46', [], ExcelValue.by_value
  attach_function 'model_c31', [], ExcelValue.by_value
  attach_function 'model_p24', [], ExcelValue.by_value
  attach_function 'model_p25', [], ExcelValue.by_value
  attach_function 'model_p23', [], ExcelValue.by_value
  attach_function 'model_p28', [], ExcelValue.by_value
  attach_function 'model_c24', [], ExcelValue.by_value
  attach_function 'model_p37', [], ExcelValue.by_value
  attach_function 'model_p38', [], ExcelValue.by_value
  attach_function 'model_p39', [], ExcelValue.by_value
  attach_function 'model_c47', [], ExcelValue.by_value
  attach_function 'model_c26', [], ExcelValue.by_value
  attach_function 'model_c25', [], ExcelValue.by_value
  attach_function 'model_c48', [], ExcelValue.by_value
  attach_function 'model_c43', [], ExcelValue.by_value
  attach_function 'model_c42', [], ExcelValue.by_value
  attach_function 'model_c41', [], ExcelValue.by_value
  attach_function 'model_c119', [], ExcelValue.by_value
  attach_function 'model_d119', [], ExcelValue.by_value
  attach_function 'model_e119', [], ExcelValue.by_value
  attach_function 'model_f119', [], ExcelValue.by_value
  attach_function 'model_g119', [], ExcelValue.by_value
  attach_function 'model_h119', [], ExcelValue.by_value
  attach_function 'model_i119', [], ExcelValue.by_value
  attach_function 'model_j119', [], ExcelValue.by_value
  attach_function 'model_k119', [], ExcelValue.by_value
  attach_function 'model_l119', [], ExcelValue.by_value
  attach_function 'model_m119', [], ExcelValue.by_value
  attach_function 'model_n119', [], ExcelValue.by_value
  attach_function 'model_o119', [], ExcelValue.by_value
  attach_function 'model_p119', [], ExcelValue.by_value
  attach_function 'model_q119', [], ExcelValue.by_value
  attach_function 'model_r119', [], ExcelValue.by_value
  attach_function 'model_s119', [], ExcelValue.by_value
  attach_function 'model_c108', [], ExcelValue.by_value
  attach_function 'model_d108', [], ExcelValue.by_value
  attach_function 'model_e108', [], ExcelValue.by_value
  attach_function 'model_f108', [], ExcelValue.by_value
  attach_function 'model_g108', [], ExcelValue.by_value
  attach_function 'model_h108', [], ExcelValue.by_value
  attach_function 'model_i108', [], ExcelValue.by_value
  attach_function 'model_j108', [], ExcelValue.by_value
  attach_function 'model_k108', [], ExcelValue.by_value
  attach_function 'model_l108', [], ExcelValue.by_value
  attach_function 'model_m108', [], ExcelValue.by_value
  attach_function 'model_n108', [], ExcelValue.by_value
  attach_function 'model_o108', [], ExcelValue.by_value
  attach_function 'model_p108', [], ExcelValue.by_value
  attach_function 'model_q108', [], ExcelValue.by_value
  attach_function 'model_r108', [], ExcelValue.by_value
  attach_function 'model_s108', [], ExcelValue.by_value
  attach_function 'model_c29', [], ExcelValue.by_value
  attach_function 'model_c110', [], ExcelValue.by_value
  attach_function 'model_d110', [], ExcelValue.by_value
  attach_function 'model_e110', [], ExcelValue.by_value
  attach_function 'model_f110', [], ExcelValue.by_value
  attach_function 'model_g110', [], ExcelValue.by_value
  attach_function 'model_h110', [], ExcelValue.by_value
  attach_function 'model_i110', [], ExcelValue.by_value
  attach_function 'model_j110', [], ExcelValue.by_value
  attach_function 'model_k110', [], ExcelValue.by_value
  attach_function 'model_l110', [], ExcelValue.by_value
  attach_function 'model_m110', [], ExcelValue.by_value
  attach_function 'model_n110', [], ExcelValue.by_value
  attach_function 'model_o110', [], ExcelValue.by_value
  attach_function 'model_p110', [], ExcelValue.by_value
  attach_function 'model_q110', [], ExcelValue.by_value
  attach_function 'model_r110', [], ExcelValue.by_value
  attach_function 'model_s110', [], ExcelValue.by_value
  attach_function 'model_c107', [], ExcelValue.by_value
  attach_function 'model_d107', [], ExcelValue.by_value
  attach_function 'model_e107', [], ExcelValue.by_value
  attach_function 'model_f107', [], ExcelValue.by_value
  attach_function 'model_g107', [], ExcelValue.by_value
  attach_function 'model_h107', [], ExcelValue.by_value
  attach_function 'model_i107', [], ExcelValue.by_value
  attach_function 'model_j107', [], ExcelValue.by_value
  attach_function 'model_k107', [], ExcelValue.by_value
  attach_function 'model_l107', [], ExcelValue.by_value
  attach_function 'model_m107', [], ExcelValue.by_value
  attach_function 'model_n107', [], ExcelValue.by_value
  attach_function 'model_o107', [], ExcelValue.by_value
  attach_function 'model_p107', [], ExcelValue.by_value
  attach_function 'model_q107', [], ExcelValue.by_value
  attach_function 'model_r107', [], ExcelValue.by_value
  attach_function 'model_s107', [], ExcelValue.by_value
  attach_function 'model_c34', [], ExcelValue.by_value
  attach_function 'model_c109', [], ExcelValue.by_value
  attach_function 'model_d109', [], ExcelValue.by_value
  attach_function 'model_e109', [], ExcelValue.by_value
  attach_function 'model_f109', [], ExcelValue.by_value
  attach_function 'model_g109', [], ExcelValue.by_value
  attach_function 'model_h109', [], ExcelValue.by_value
  attach_function 'model_i109', [], ExcelValue.by_value
  attach_function 'model_j109', [], ExcelValue.by_value
  attach_function 'model_k109', [], ExcelValue.by_value
  attach_function 'model_l109', [], ExcelValue.by_value
  attach_function 'model_m109', [], ExcelValue.by_value
  attach_function 'model_n109', [], ExcelValue.by_value
  attach_function 'model_o109', [], ExcelValue.by_value
  attach_function 'model_p109', [], ExcelValue.by_value
  attach_function 'model_q109', [], ExcelValue.by_value
  attach_function 'model_r109', [], ExcelValue.by_value
  attach_function 'model_s109', [], ExcelValue.by_value
  attach_function 'model_c114', [], ExcelValue.by_value
  attach_function 'model_d114', [], ExcelValue.by_value
  attach_function 'model_e114', [], ExcelValue.by_value
  attach_function 'model_f114', [], ExcelValue.by_value
  attach_function 'model_g114', [], ExcelValue.by_value
  attach_function 'model_h114', [], ExcelValue.by_value
  attach_function 'model_i114', [], ExcelValue.by_value
  attach_function 'model_j114', [], ExcelValue.by_value
  attach_function 'model_k114', [], ExcelValue.by_value
  attach_function 'model_l114', [], ExcelValue.by_value
  attach_function 'model_m114', [], ExcelValue.by_value
  attach_function 'model_n114', [], ExcelValue.by_value
  attach_function 'model_o114', [], ExcelValue.by_value
  attach_function 'model_p114', [], ExcelValue.by_value
  attach_function 'model_q114', [], ExcelValue.by_value
  attach_function 'model_r114', [], ExcelValue.by_value
  attach_function 'model_s114', [], ExcelValue.by_value
  attach_function 'model_c116', [], ExcelValue.by_value
  attach_function 'model_d116', [], ExcelValue.by_value
  attach_function 'model_e116', [], ExcelValue.by_value
  attach_function 'model_f116', [], ExcelValue.by_value
  attach_function 'model_g116', [], ExcelValue.by_value
  attach_function 'model_h116', [], ExcelValue.by_value
  attach_function 'model_i116', [], ExcelValue.by_value
  attach_function 'model_j116', [], ExcelValue.by_value
  attach_function 'model_k116', [], ExcelValue.by_value
  attach_function 'model_l116', [], ExcelValue.by_value
  attach_function 'model_m116', [], ExcelValue.by_value
  attach_function 'model_n116', [], ExcelValue.by_value
  attach_function 'model_o116', [], ExcelValue.by_value
  attach_function 'model_p116', [], ExcelValue.by_value
  attach_function 'model_q116', [], ExcelValue.by_value
  attach_function 'model_r116', [], ExcelValue.by_value
  attach_function 'model_s116', [], ExcelValue.by_value
  attach_function 'model_c113', [], ExcelValue.by_value
  attach_function 'model_d113', [], ExcelValue.by_value
  attach_function 'model_e113', [], ExcelValue.by_value
  attach_function 'model_f113', [], ExcelValue.by_value
  attach_function 'model_g113', [], ExcelValue.by_value
  attach_function 'model_h113', [], ExcelValue.by_value
  attach_function 'model_i113', [], ExcelValue.by_value
  attach_function 'model_j113', [], ExcelValue.by_value
  attach_function 'model_k113', [], ExcelValue.by_value
  attach_function 'model_l113', [], ExcelValue.by_value
  attach_function 'model_m113', [], ExcelValue.by_value
  attach_function 'model_n113', [], ExcelValue.by_value
  attach_function 'model_o113', [], ExcelValue.by_value
  attach_function 'model_p113', [], ExcelValue.by_value
  attach_function 'model_q113', [], ExcelValue.by_value
  attach_function 'model_r113', [], ExcelValue.by_value
  attach_function 'model_s113', [], ExcelValue.by_value
  attach_function 'model_c117', [], ExcelValue.by_value
  attach_function 'model_d117', [], ExcelValue.by_value
  attach_function 'model_e117', [], ExcelValue.by_value
  attach_function 'model_f117', [], ExcelValue.by_value
  attach_function 'model_g117', [], ExcelValue.by_value
  attach_function 'model_h117', [], ExcelValue.by_value
  attach_function 'model_i117', [], ExcelValue.by_value
  attach_function 'model_j117', [], ExcelValue.by_value
  attach_function 'model_k117', [], ExcelValue.by_value
  attach_function 'model_l117', [], ExcelValue.by_value
  attach_function 'model_m117', [], ExcelValue.by_value
  attach_function 'model_n117', [], ExcelValue.by_value
  attach_function 'model_o117', [], ExcelValue.by_value
  attach_function 'model_p117', [], ExcelValue.by_value
  attach_function 'model_q117', [], ExcelValue.by_value
  attach_function 'model_r117', [], ExcelValue.by_value
  attach_function 'model_s117', [], ExcelValue.by_value
  attach_function 'model_c115', [], ExcelValue.by_value
  attach_function 'model_d115', [], ExcelValue.by_value
  attach_function 'model_e115', [], ExcelValue.by_value
  attach_function 'model_f115', [], ExcelValue.by_value
  attach_function 'model_g115', [], ExcelValue.by_value
  attach_function 'model_h115', [], ExcelValue.by_value
  attach_function 'model_i115', [], ExcelValue.by_value
  attach_function 'model_j115', [], ExcelValue.by_value
  attach_function 'model_k115', [], ExcelValue.by_value
  attach_function 'model_l115', [], ExcelValue.by_value
  attach_function 'model_m115', [], ExcelValue.by_value
  attach_function 'model_n115', [], ExcelValue.by_value
  attach_function 'model_o115', [], ExcelValue.by_value
  attach_function 'model_p115', [], ExcelValue.by_value
  attach_function 'model_q115', [], ExcelValue.by_value
  attach_function 'model_r115', [], ExcelValue.by_value
  attach_function 'model_s115', [], ExcelValue.by_value
  attach_function 'model_r31', [], ExcelValue.by_value
  attach_function 'model_q31', [], ExcelValue.by_value
  attach_function 'model_p31', [], ExcelValue.by_value
  attach_function 'model_r32', [], ExcelValue.by_value
  attach_function 'model_q32', [], ExcelValue.by_value
  attach_function 'model_p32', [], ExcelValue.by_value
  attach_function 'model_r33', [], ExcelValue.by_value
  attach_function 'model_q33', [], ExcelValue.by_value
  attach_function 'model_p33', [], ExcelValue.by_value
  attach_function 'model_r34', [], ExcelValue.by_value
  attach_function 'model_q34', [], ExcelValue.by_value
  attach_function 'model_p34', [], ExcelValue.by_value
  attach_function 'model_c139', [], ExcelValue.by_value
  attach_function 'model_d139', [], ExcelValue.by_value
  attach_function 'model_e139', [], ExcelValue.by_value
  attach_function 'model_f139', [], ExcelValue.by_value
  attach_function 'model_g139', [], ExcelValue.by_value
  attach_function 'model_h139', [], ExcelValue.by_value
  attach_function 'model_i139', [], ExcelValue.by_value
  attach_function 'model_j139', [], ExcelValue.by_value
  attach_function 'model_k139', [], ExcelValue.by_value
  attach_function 'model_l139', [], ExcelValue.by_value
  attach_function 'model_m139', [], ExcelValue.by_value
  attach_function 'model_n139', [], ExcelValue.by_value
  attach_function 'model_o139', [], ExcelValue.by_value
  attach_function 'model_p139', [], ExcelValue.by_value
  attach_function 'model_q139', [], ExcelValue.by_value
  attach_function 'model_r139', [], ExcelValue.by_value
  attach_function 'model_s139', [], ExcelValue.by_value
  attach_function 'model_c137', [], ExcelValue.by_value
  attach_function 'model_d137', [], ExcelValue.by_value
  attach_function 'model_e137', [], ExcelValue.by_value
  attach_function 'model_f137', [], ExcelValue.by_value
  attach_function 'model_g137', [], ExcelValue.by_value
  attach_function 'model_h137', [], ExcelValue.by_value
  attach_function 'model_i137', [], ExcelValue.by_value
  attach_function 'model_j137', [], ExcelValue.by_value
  attach_function 'model_k137', [], ExcelValue.by_value
  attach_function 'model_l137', [], ExcelValue.by_value
  attach_function 'model_m137', [], ExcelValue.by_value
  attach_function 'model_n137', [], ExcelValue.by_value
  attach_function 'model_o137', [], ExcelValue.by_value
  attach_function 'model_p137', [], ExcelValue.by_value
  attach_function 'model_q137', [], ExcelValue.by_value
  attach_function 'model_r137', [], ExcelValue.by_value
  attach_function 'model_s137', [], ExcelValue.by_value
  attach_function 'model_c140', [], ExcelValue.by_value
  attach_function 'model_d140', [], ExcelValue.by_value
  attach_function 'model_e140', [], ExcelValue.by_value
  attach_function 'model_f140', [], ExcelValue.by_value
  attach_function 'model_g140', [], ExcelValue.by_value
  attach_function 'model_h140', [], ExcelValue.by_value
  attach_function 'model_i140', [], ExcelValue.by_value
  attach_function 'model_j140', [], ExcelValue.by_value
  attach_function 'model_k140', [], ExcelValue.by_value
  attach_function 'model_l140', [], ExcelValue.by_value
  attach_function 'model_m140', [], ExcelValue.by_value
  attach_function 'model_n140', [], ExcelValue.by_value
  attach_function 'model_o140', [], ExcelValue.by_value
  attach_function 'model_p140', [], ExcelValue.by_value
  attach_function 'model_q140', [], ExcelValue.by_value
  attach_function 'model_r140', [], ExcelValue.by_value
  attach_function 'model_s140', [], ExcelValue.by_value
  attach_function 'model_c141', [], ExcelValue.by_value
  attach_function 'model_d141', [], ExcelValue.by_value
  attach_function 'model_e141', [], ExcelValue.by_value
  attach_function 'model_f141', [], ExcelValue.by_value
  attach_function 'model_g141', [], ExcelValue.by_value
  attach_function 'model_h141', [], ExcelValue.by_value
  attach_function 'model_i141', [], ExcelValue.by_value
  attach_function 'model_j141', [], ExcelValue.by_value
  attach_function 'model_k141', [], ExcelValue.by_value
  attach_function 'model_l141', [], ExcelValue.by_value
  attach_function 'model_m141', [], ExcelValue.by_value
  attach_function 'model_n141', [], ExcelValue.by_value
  attach_function 'model_o141', [], ExcelValue.by_value
  attach_function 'model_p141', [], ExcelValue.by_value
  attach_function 'model_q141', [], ExcelValue.by_value
  attach_function 'model_r141', [], ExcelValue.by_value
  attach_function 'model_s141', [], ExcelValue.by_value
  attach_function 'model_c143', [], ExcelValue.by_value
  attach_function 'model_d143', [], ExcelValue.by_value
  attach_function 'model_e143', [], ExcelValue.by_value
  attach_function 'model_f143', [], ExcelValue.by_value
  attach_function 'model_g143', [], ExcelValue.by_value
  attach_function 'model_h143', [], ExcelValue.by_value
  attach_function 'model_i143', [], ExcelValue.by_value
  attach_function 'model_j143', [], ExcelValue.by_value
  attach_function 'model_k143', [], ExcelValue.by_value
  attach_function 'model_l143', [], ExcelValue.by_value
  attach_function 'model_m143', [], ExcelValue.by_value
  attach_function 'model_n143', [], ExcelValue.by_value
  attach_function 'model_o143', [], ExcelValue.by_value
  attach_function 'model_p143', [], ExcelValue.by_value
  attach_function 'model_q143', [], ExcelValue.by_value
  attach_function 'model_r143', [], ExcelValue.by_value
  attach_function 'model_s143', [], ExcelValue.by_value
  attach_function 'model_c142', [], ExcelValue.by_value
  attach_function 'model_d142', [], ExcelValue.by_value
  attach_function 'model_e142', [], ExcelValue.by_value
  attach_function 'model_f142', [], ExcelValue.by_value
  attach_function 'model_g142', [], ExcelValue.by_value
  attach_function 'model_h142', [], ExcelValue.by_value
  attach_function 'model_i142', [], ExcelValue.by_value
  attach_function 'model_j142', [], ExcelValue.by_value
  attach_function 'model_k142', [], ExcelValue.by_value
  attach_function 'model_l142', [], ExcelValue.by_value
  attach_function 'model_m142', [], ExcelValue.by_value
  attach_function 'model_n142', [], ExcelValue.by_value
  attach_function 'model_o142', [], ExcelValue.by_value
  attach_function 'model_p142', [], ExcelValue.by_value
  attach_function 'model_q142', [], ExcelValue.by_value
  attach_function 'model_r142', [], ExcelValue.by_value
  attach_function 'model_s142', [], ExcelValue.by_value
  attach_function 'model_c132', [], ExcelValue.by_value
  attach_function 'model_d132', [], ExcelValue.by_value
  attach_function 'model_e132', [], ExcelValue.by_value
  attach_function 'model_f132', [], ExcelValue.by_value
  attach_function 'model_g132', [], ExcelValue.by_value
  attach_function 'model_h132', [], ExcelValue.by_value
  attach_function 'model_i132', [], ExcelValue.by_value
  attach_function 'model_j132', [], ExcelValue.by_value
  attach_function 'model_k132', [], ExcelValue.by_value
  attach_function 'model_l132', [], ExcelValue.by_value
  attach_function 'model_m132', [], ExcelValue.by_value
  attach_function 'model_n132', [], ExcelValue.by_value
  attach_function 'model_o132', [], ExcelValue.by_value
  attach_function 'model_p132', [], ExcelValue.by_value
  attach_function 'model_q132', [], ExcelValue.by_value
  attach_function 'model_r132', [], ExcelValue.by_value
  attach_function 'model_s132', [], ExcelValue.by_value
  attach_function 'model_c1252', [], ExcelValue.by_value
  attach_function 'model_c1250', [], ExcelValue.by_value
  attach_function 'model_c1248', [], ExcelValue.by_value
  attach_function 'model_c131', [], ExcelValue.by_value
  attach_function 'model_d131', [], ExcelValue.by_value
  attach_function 'model_e131', [], ExcelValue.by_value
  attach_function 'model_f131', [], ExcelValue.by_value
  attach_function 'model_g131', [], ExcelValue.by_value
  attach_function 'model_h131', [], ExcelValue.by_value
  attach_function 'model_i131', [], ExcelValue.by_value
  attach_function 'model_j131', [], ExcelValue.by_value
  attach_function 'model_k131', [], ExcelValue.by_value
  attach_function 'model_l131', [], ExcelValue.by_value
  attach_function 'model_m131', [], ExcelValue.by_value
  attach_function 'model_n131', [], ExcelValue.by_value
  attach_function 'model_o131', [], ExcelValue.by_value
  attach_function 'model_p131', [], ExcelValue.by_value
  attach_function 'model_q131', [], ExcelValue.by_value
  attach_function 'model_r131', [], ExcelValue.by_value
  attach_function 'model_s131', [], ExcelValue.by_value
  attach_function 'model_c1249', [], ExcelValue.by_value
  attach_function 'model_c1251', [], ExcelValue.by_value
  attach_function 'model_c133', [], ExcelValue.by_value
  attach_function 'model_d133', [], ExcelValue.by_value
  attach_function 'model_e133', [], ExcelValue.by_value
  attach_function 'model_f133', [], ExcelValue.by_value
  attach_function 'model_g133', [], ExcelValue.by_value
  attach_function 'model_h133', [], ExcelValue.by_value
  attach_function 'model_i133', [], ExcelValue.by_value
  attach_function 'model_j133', [], ExcelValue.by_value
  attach_function 'model_k133', [], ExcelValue.by_value
  attach_function 'model_l133', [], ExcelValue.by_value
  attach_function 'model_m133', [], ExcelValue.by_value
  attach_function 'model_n133', [], ExcelValue.by_value
  attach_function 'model_o133', [], ExcelValue.by_value
  attach_function 'model_p133', [], ExcelValue.by_value
  attach_function 'model_q133', [], ExcelValue.by_value
  attach_function 'model_r133', [], ExcelValue.by_value
  attach_function 'model_s133', [], ExcelValue.by_value
  attach_function 'model_c77', [], ExcelValue.by_value
  attach_function 'model_c126', [], ExcelValue.by_value
  attach_function 'model_d126', [], ExcelValue.by_value
  attach_function 'model_e126', [], ExcelValue.by_value
  attach_function 'model_f126', [], ExcelValue.by_value
  attach_function 'model_g126', [], ExcelValue.by_value
  attach_function 'model_h126', [], ExcelValue.by_value
  attach_function 'model_i126', [], ExcelValue.by_value
  attach_function 'model_j126', [], ExcelValue.by_value
  attach_function 'model_k126', [], ExcelValue.by_value
  attach_function 'model_l126', [], ExcelValue.by_value
  attach_function 'model_m126', [], ExcelValue.by_value
  attach_function 'model_n126', [], ExcelValue.by_value
  attach_function 'model_o126', [], ExcelValue.by_value
  attach_function 'model_p126', [], ExcelValue.by_value
  attach_function 'model_q126', [], ExcelValue.by_value
  attach_function 'model_r126', [], ExcelValue.by_value
  attach_function 'model_s126', [], ExcelValue.by_value
  attach_function 'model_c128', [], ExcelValue.by_value
  attach_function 'model_d128', [], ExcelValue.by_value
  attach_function 'model_e128', [], ExcelValue.by_value
  attach_function 'model_f128', [], ExcelValue.by_value
  attach_function 'model_g128', [], ExcelValue.by_value
  attach_function 'model_h128', [], ExcelValue.by_value
  attach_function 'model_i128', [], ExcelValue.by_value
  attach_function 'model_j128', [], ExcelValue.by_value
  attach_function 'model_k128', [], ExcelValue.by_value
  attach_function 'model_l128', [], ExcelValue.by_value
  attach_function 'model_m128', [], ExcelValue.by_value
  attach_function 'model_n128', [], ExcelValue.by_value
  attach_function 'model_o128', [], ExcelValue.by_value
  attach_function 'model_p128', [], ExcelValue.by_value
  attach_function 'model_q128', [], ExcelValue.by_value
  attach_function 'model_r128', [], ExcelValue.by_value
  attach_function 'model_s128', [], ExcelValue.by_value
  attach_function 'model_c1244', [], ExcelValue.by_value
  attach_function 'model_c1245', [], ExcelValue.by_value
  attach_function 'model_c1243', [], ExcelValue.by_value
  attach_function 'model_c124', [], ExcelValue.by_value
  attach_function 'model_d124', [], ExcelValue.by_value
  attach_function 'model_e124', [], ExcelValue.by_value
  attach_function 'model_f124', [], ExcelValue.by_value
  attach_function 'model_g124', [], ExcelValue.by_value
  attach_function 'model_h124', [], ExcelValue.by_value
  attach_function 'model_i124', [], ExcelValue.by_value
  attach_function 'model_j124', [], ExcelValue.by_value
  attach_function 'model_k124', [], ExcelValue.by_value
  attach_function 'model_l124', [], ExcelValue.by_value
  attach_function 'model_m124', [], ExcelValue.by_value
  attach_function 'model_n124', [], ExcelValue.by_value
  attach_function 'model_o124', [], ExcelValue.by_value
  attach_function 'model_p124', [], ExcelValue.by_value
  attach_function 'model_q124', [], ExcelValue.by_value
  attach_function 'model_r124', [], ExcelValue.by_value
  attach_function 'model_s124', [], ExcelValue.by_value
  attach_function 'model_c39', [], ExcelValue.by_value
  attach_function 'model_c38', [], ExcelValue.by_value
  attach_function 'model_c120', [], ExcelValue.by_value
  attach_function 'model_d120', [], ExcelValue.by_value
  attach_function 'model_e120', [], ExcelValue.by_value
  attach_function 'model_f120', [], ExcelValue.by_value
  attach_function 'model_g120', [], ExcelValue.by_value
  attach_function 'model_h120', [], ExcelValue.by_value
  attach_function 'model_i120', [], ExcelValue.by_value
  attach_function 'model_j120', [], ExcelValue.by_value
  attach_function 'model_k120', [], ExcelValue.by_value
  attach_function 'model_l120', [], ExcelValue.by_value
  attach_function 'model_m120', [], ExcelValue.by_value
  attach_function 'model_n120', [], ExcelValue.by_value
  attach_function 'model_o120', [], ExcelValue.by_value
  attach_function 'model_p120', [], ExcelValue.by_value
  attach_function 'model_q120', [], ExcelValue.by_value
  attach_function 'model_r120', [], ExcelValue.by_value
  attach_function 'model_s120', [], ExcelValue.by_value
  attach_function 'model_c37', [], ExcelValue.by_value
  attach_function 'model_c121', [], ExcelValue.by_value
  attach_function 'model_d121', [], ExcelValue.by_value
  attach_function 'model_e121', [], ExcelValue.by_value
  attach_function 'model_f121', [], ExcelValue.by_value
  attach_function 'model_g121', [], ExcelValue.by_value
  attach_function 'model_h121', [], ExcelValue.by_value
  attach_function 'model_i121', [], ExcelValue.by_value
  attach_function 'model_j121', [], ExcelValue.by_value
  attach_function 'model_k121', [], ExcelValue.by_value
  attach_function 'model_l121', [], ExcelValue.by_value
  attach_function 'model_m121', [], ExcelValue.by_value
  attach_function 'model_n121', [], ExcelValue.by_value
  attach_function 'model_o121', [], ExcelValue.by_value
  attach_function 'model_p121', [], ExcelValue.by_value
  attach_function 'model_q121', [], ExcelValue.by_value
  attach_function 'model_r121', [], ExcelValue.by_value
  attach_function 'model_s121', [], ExcelValue.by_value
  attach_function 'model_c30', [], ExcelValue.by_value
  attach_function 'model_c35', [], ExcelValue.by_value
  attach_function 'model_c76', [], ExcelValue.by_value
  attach_function 'model_c23', [], ExcelValue.by_value
  attach_function 'model_c33', [], ExcelValue.by_value
  # end of Model
  # Start of named references
  attach_function 'n_2012_onwards_electricity_demand_growth_rate', [], ExcelValue.by_value
  attach_function 'n_2020_fossil_fuel_emissions_factor', [], ExcelValue.by_value
  attach_function 'n_2020_non_renewable_low_carbon_generation_i_e_nuclear_ccs', [], ExcelValue.by_value
  attach_function 'n_2020_renewables_target', [], ExcelValue.by_value
  attach_function 'n_2020_renewables_target_gco2_kwh', [], ExcelValue.by_value
  attach_function 'n_2020_renewables_target_twh', [], ExcelValue.by_value
  attach_function 'n_2030_decarbonisation_level', [], ExcelValue.by_value
  attach_function 'n_2050_electricity_demand', [], ExcelValue.by_value
  attach_function 'n_2050_emissions_electricity', [], ExcelValue.by_value
  attach_function 'n_2050_emissions_industry', [], ExcelValue.by_value
  attach_function 'n_2050_emissions_total', [], ExcelValue.by_value
  attach_function 'n_2050_fossil_fuel_emissions_factor', [], ExcelValue.by_value
  attach_function 'n_2050_maximum_electricity_demand', [], ExcelValue.by_value
  attach_function 'n_2050_minimum_electricity_demand', [], ExcelValue.by_value
  attach_function 'annual_change_in_non_electricity_traded_emissions', [], ExcelValue.by_value
  attach_function 'average_life_high_carbon', [], ExcelValue.by_value
  attach_function 'average_life_other_low_carbon', [], ExcelValue.by_value
  attach_function 'average_life_wind', [], ExcelValue.by_value
  attach_function 'baseload_demand', [], ExcelValue.by_value
  attach_function 'build_rate_dispatchable_low_carbon', [], ExcelValue.by_value
  attach_function 'build_rate_from_now_to_2020', [], ExcelValue.by_value
  attach_function 'build_rate_high_carbon', [], ExcelValue.by_value
  attach_function 'build_rate_intermittent_low_carbon', [], ExcelValue.by_value
  attach_function 'build_rate_target_in_second_build', [], ExcelValue.by_value
  attach_function 'build_rate_total_low_carbon', [], ExcelValue.by_value
  attach_function 'capacity_dispatchable_low_carbon', [], ExcelValue.by_value
  attach_function 'capacity_high_carbon', [], ExcelValue.by_value
  attach_function 'capacity_intermittent_low_carbon', [], ExcelValue.by_value
  attach_function 'capacity_total', [], ExcelValue.by_value
  attach_function 'capacity_total_low_carbon', [], ExcelValue.by_value
  attach_function 'cb2_net_ets_purchase', [], ExcelValue.by_value
  attach_function 'cb2_scenario', [], ExcelValue.by_value
  attach_function 'cb2_traded_cap', [], ExcelValue.by_value
  attach_function 'cb3_net_ets_purchase', [], ExcelValue.by_value
  attach_function 'cb3_scenario', [], ExcelValue.by_value
  attach_function 'cb3_traded_cap', [], ExcelValue.by_value
  attach_function 'cb4_current_net_ets_purchase', [], ExcelValue.by_value
  attach_function 'cb4_current_scenario', [], ExcelValue.by_value
  attach_function 'cb4_current_traded_cap', [], ExcelValue.by_value
  attach_function 'cb4_revised_net_ets_purchase', [], ExcelValue.by_value
  attach_function 'cb4_revised_scenario', [], ExcelValue.by_value
  attach_function 'cb4_revised_traded_cap', [], ExcelValue.by_value
  attach_function 'emissions_electicity', [], ExcelValue.by_value
  attach_function 'emissions_factor', [], ExcelValue.by_value
  attach_function 'emissions_non_electricity_traded', [], ExcelValue.by_value
  attach_function 'emissions_total_traded', [], ExcelValue.by_value
  attach_function 'emissions_uk_share_of_eu_ets_cap_alternative', [], ExcelValue.by_value
  attach_function 'emissions_uk_share_of_eu_ets_cap_current', [], ExcelValue.by_value
  attach_function 'energy_output_dispatchable_low_carbon', [], ExcelValue.by_value
  attach_function 'energy_output_error', [], ExcelValue.by_value
  attach_function 'energy_output_high_carbon', [], ExcelValue.by_value
  attach_function 'energy_output_intermittent', [], ExcelValue.by_value
  attach_function 'energy_output_intermittent_low_carbon', [], ExcelValue.by_value
  attach_function 'energy_output_low_carbon', [], ExcelValue.by_value
  attach_function 'energy_output_total', [], ExcelValue.by_value
  attach_function 'energy_output_total_low_carbon', [], ExcelValue.by_value
  attach_function 'gw_per_twh', [], ExcelValue.by_value
  attach_function 'load_factor_average_low_carbon', [], ExcelValue.by_value
  attach_function 'load_factor_demand', [], ExcelValue.by_value
  attach_function 'load_factor_dispatchable_low_carbon', [], ExcelValue.by_value
  attach_function 'load_factor_high_carbon', [], ExcelValue.by_value
  attach_function 'load_factor_intermittent', [], ExcelValue.by_value
  attach_function 'load_factor_intermittent_low_carbon', [], ExcelValue.by_value
  attach_function 'maximum_industry_contraction', [], ExcelValue.by_value
  attach_function 'maximum_industry_expansion', [], ExcelValue.by_value
  attach_function 'mean_demand', [], ExcelValue.by_value
  attach_function 'minimum_build_rate', [], ExcelValue.by_value
  attach_function 'peak_demand', [], ExcelValue.by_value
  attach_function 'proportion_of_build_rate_to_2020_that_is_wind_rest_is_bio', [], ExcelValue.by_value
  attach_function 'proportion_of_second_build_that_is_wind', [], ExcelValue.by_value
  attach_function 'twh_per_gw', [], ExcelValue.by_value
  attach_function 'year_electricity_demand_starts_to_increase', [], ExcelValue.by_value
  attach_function 'year_second_wave_of_building_starts', [], ExcelValue.by_value
  attach_function 'set_n_2012_onwards_electricity_demand_growth_rate', [ExcelValue.by_value], :void
  attach_function 'set_n_2020_fossil_fuel_emissions_factor', [ExcelValue.by_value], :void
  attach_function 'set_n_2050_electricity_demand', [ExcelValue.by_value], :void
  attach_function 'set_n_2050_fossil_fuel_emissions_factor', [ExcelValue.by_value], :void
  attach_function 'set_n_2050_maximum_electricity_demand', [ExcelValue.by_value], :void
  attach_function 'set_n_2050_minimum_electricity_demand', [ExcelValue.by_value], :void
  attach_function 'set_annual_change_in_non_electricity_traded_emissions', [ExcelValue.by_value], :void
  attach_function 'set_average_life_high_carbon', [ExcelValue.by_value], :void
  attach_function 'set_average_life_other_low_carbon', [ExcelValue.by_value], :void
  attach_function 'set_average_life_wind', [ExcelValue.by_value], :void
  attach_function 'set_build_rate_from_now_to_2020', [ExcelValue.by_value], :void
  attach_function 'set_build_rate_target_in_second_build', [ExcelValue.by_value], :void
  attach_function 'set_maximum_industry_contraction', [ExcelValue.by_value], :void
  attach_function 'set_maximum_industry_expansion', [ExcelValue.by_value], :void
  attach_function 'set_minimum_build_rate', [ExcelValue.by_value], :void
  attach_function 'set_proportion_of_build_rate_to_2020_that_is_wind_rest_is_bio', [ExcelValue.by_value], :void
  attach_function 'set_proportion_of_second_build_that_is_wind', [ExcelValue.by_value], :void
  attach_function 'set_year_electricity_demand_starts_to_increase', [ExcelValue.by_value], :void
  attach_function 'set_year_second_wave_of_building_starts', [ExcelValue.by_value], :void
  # End of named references
end
