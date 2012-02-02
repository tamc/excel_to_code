require_relative '../spec_helper'

describe ExpandArrayFormulaeAst do
    
  before(:all) do 
    @mapper = ExpandArrayFormulaeAst.new
  end
  
  checks = test_data('expand_array_formulae_ast.txt').lines.map.with_index { |line,i| [i,line] }.find_all { |line| line[1].start_with?('[:') }.map { |line| [line[0],*line[1].split("\t")] }
  checks.each do |c|
    it "converts array formula #{c[1].strip} into #{c[4].strip}" do
      desired = c[4].strip
      actual = @mapper.map(eval(c[1]))
      actual.to_s.should == desired.to_s
    end
  end
end
