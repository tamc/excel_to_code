#!/bin/env ruby

unless ARGV.length == 3
  puts "Usage: #{$0} format c_filename function_name"
  puts
  puts "format = names - print the function and constant names this function depends on"
  puts "format = bodies - print the function bodies and constant values this function depends on"
  puts "format = values - print out a c file that, when run, will print the values of all the functions and constants"
  exit
end

$format, cfile, function_name = *ARGV
function_name += "()"

f = IO.readlines(cfile).join

def function(f, name) 
  f[/^(static )?ExcelValue #{Regexp.escape(name)} {.*?^}/m]
end

def constant(f, name)
  f[/^(static )?ExcelValue #{Regexp.escape(name)} = {.*?};/m]
end

def constants(function_body)
  function_body.scan(/\bconstant\d+\b/m)
end

def dependencies(function_body)
  function_body.scan(/\b[a-z0-9_]+_[a-z]+\d+\(\)/m)+function_body.scan(/\bcommon\d+\(\)/)
end

def recursive(f, function_name, seen = {})
  seen[function_name] = true
  fb = function(f, function_name)
  case $format 
  when 'names' 
    puts function_name
  when 'bodies'
    puts fb
  when 'values'
    puts "printf(\"#{function_name} \"); inspect_excel_value(#{function_name});"
  end
  constants(fb).each do |c|
    next if seen[c]
    seen[c] = true

    cb = constant(f, c)
    case $format 
    when 'names'
      puts c
    when 'bodies'
      puts cb
    when 'values'
      puts "printf(\"#{c} \"); inspect_excel_value(#{c});"
    end
  end
  dependencies(fb).each do |fn|
    recursive(f, fn, seen) unless seen[fn]
  end
end

if $format == "values"
  puts "#include <stdio.h>"
  puts "#include \"#{cfile}\""
  puts "int main() {"
end

recursive(f, function_name)

if $format == "values"
  puts "return 0;"
  puts "}"
end
