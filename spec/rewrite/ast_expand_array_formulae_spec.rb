require_relative '../spec_helper'

describe AstExpandArrayFormulae do
    
  before(:all) do 
    @mapper = AstExpandArrayFormulae.new
  end
  
  checks = test_data('ast_expand_array_formulae.txt').each_line.map.with_index { |line,i| [i,line] }.find_all { |line| line[1].start_with?('[:') }.map { |line| [line[0],*line[1].split("\t")] }
  checks.each do |c|
    it "converts array formula #{c[1].strip} into #{c[2].strip}" do
      desired = eval(c[2].strip)
      actual = @mapper.map(eval(c[1]))
      actual.should == desired
    end
  end
end
