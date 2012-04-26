require 'ffi'

module Blank
  extend FFI::Library
  ffi_lib  File.join(File.dirname(__FILE__),FFI.map_library_name('blank'))
  ExcelType = enum :ExcelEmpty, :ExcelNumber, :ExcelString, :ExcelBoolean, :ExcelError, :ExcelRange
                
  class ExcelValue < FFI::Struct
    layout :type, ExcelType,
  	       :number, :double,
  	       :string, :string,
         	 :array, :pointer,
           :rows, :int,
           :columns, :int             
  end
  

  # use this function to reset all cell values
  attach_function 'reset', [], :void

  # start of Sheet1
  attach_function 'set_sheet1_a1', [ExcelValue.by_value], :void
  attach_function 'sheet1_a2', [], ExcelValue.by_value
  # end of Sheet1
end
