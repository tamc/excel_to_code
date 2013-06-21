require_relative '../spec_helper'
require 'textpeg2rubypeg'

describe Formula do
  
  before(:all) do
    text_peg = File.join(File.dirname(__FILE__),'..','..','src','excel','formula_peg.txt')
    ruby_peg = File.join(File.dirname(__FILE__),'..','..','src','excel','formula_peg.rb') 
    ast = TextPeg.parse(IO.readlines(text_peg).join)
    builder = TextPeg2RubyPeg.new
    new_ruby = ast.visit(builder)
    
    old_ruby = IO.readlines(ruby_peg).join
    
    if new_ruby != old_ruby
      puts "Excel formulae parser updated"    
      File.open(ruby_peg,'w') do |f|
        f.puts new_ruby
      end
      Kernel.eval(new_ruby)
    end  
  end
  
  checks = test_data('formulae_to_ast.txt').each_line.map.with_index { |line,i| [i,line] }.find_all { |line| line.last =~ /\[:/ }.map { |line| line.last =~ /(.*?)(\[:.*)/; [line.first,$1,$2] }
  checks.each do |c|
    it "converts #{c[1].strip} into #{c[2].strip} (line #{c[0]+1} of formulae_to_ast.txt)" do
      desired = eval(c[2].strip)
      parser = Formula.new      
      actual = parser.parse(c[1].strip)
      if actual
        actual = actual.to_ast[1]
      else
        actual = "Failed to parse"
      end
      actual.to_s.should == desired.to_s
      unless actual.to_s == desired.to_s
        parser.pretty_print_cache(true)
      end
    end
  end
end
