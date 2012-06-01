require 'ffi'
require 'fileutils'
include FileUtils

this_directory = File.dirname(__FILE__)

rm_f "Makefile"

%w{blank eu}.each do |name|
  rm_f File.join(this_directory,name,".c")
  rm_f File.join(this_directory,name,".o")
  rm_f File.join(this_directory,FFI.map_library_name(name))
  rm_f File.join(this_directory,name,".rb")
  rm_f File.join(this_directory,"test_#{name}",".rb")
  rm_rf File.join(this_directory,name)
end
  