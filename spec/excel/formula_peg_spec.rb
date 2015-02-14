require_relative '../spec_helper'
require 'textpeg2rubypeg'

describe Formula do

  # The parser is written as formula_peg.txt and then compiled into formula_peg.rb
  # This method checks whether formula_peg.txt has been updated, and if so, updated
  # formula_peg.rb 
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
  
  # The test data is stored in formulae_to_ast.txt in the format:
  # Formula as text <tab> Expected ast for formula <newline>
  # Anything that doesn't look like that is skipped.
  checks = test_data('formulae_to_ast.txt').each_line.map.with_index do |line,i| 
    [i,line]
  end.find_all do |line|
    line.last =~ /\[:/ 
  end.map do |line|
    line.last =~ /(.*?)(\[:.*)/
    [line.first,$1,$2]
  end

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
