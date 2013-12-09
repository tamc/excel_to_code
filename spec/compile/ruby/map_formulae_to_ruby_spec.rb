require_relative '../../spec_helper'

describe MapFormulaeToRuby do
    
  before(:all) do 
    @mapper = MapFormulaeToRuby.new
    @mapper.sheet_names = {'test' => 'test', 'A complicated sheet name' => 'a_complicated_sheet_name'}
  end
  
  checks = test_data('formulae_to_ruby.txt').each_line.map.with_index { |line,i| [i,line] }.find_all { |line| line[1].start_with?('[:') }.map { |line| [line[0],*line[1].split("\t")] }
  checks.each do |c|
    it "converts #{c[1].strip} into #{c[2].strip} (line #{c[0]+1} of formulae_to_ruby.txt)" do
      desired = c[2].strip
      actual = @mapper.map(eval(c[1]))
      actual.to_s.should == desired.to_s
    end
  end
  
end
