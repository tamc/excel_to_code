require_relative '../../src/excel_to_code'

this_directory = File.dirname(__FILE__)

input = File.open(File.join(this_directory,'formulae.ast'))
common = File.open(File.join(this_directory,'common-elements.ast'))
output = StringIO.new
expected_output = IO.readlines(File.join(this_directory,'formulae-expected.ast')).join

start_time = Time.now
r = ReplaceCommonElementsInFormulae.new
r.replace(input,common,output)
end_time = Time.now

input.close
common.close

unless output.string == expected_output
  puts "Wrong results"
  File.open("formulae-actual.ast",'w') do |f|
    f.puts output.string
  end
end

puts "Elapsed time #{end_time - start_time}s"
