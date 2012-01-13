require_relative '../spec_helper'

describe RewriteArrayFormulaAst do
    
  before(:all) do 
    @mapper = RewriteArrayFormulaAst.new
  end
  
  checks = test_data('array_formulae_expansion.txt').lines.map.with_index { |line,i| [i,line] }.find_all { |line| line[1].start_with?('[:') }.map { |line| [line[0],*line[1].split("\t")] }
  checks.each do |c|
    it "converts array formula #{c[1].strip} into #{c[4].strip} when offseting by #{c[2]},#{c[3]} (line #{c[0]+1} of formulae_to_ast.txt)" do
      desired = c[4].strip
      @mapper.row_offset = c[2].to_i
      @mapper.column_offset = c[3].to_i
      actual = @mapper.map(eval(c[1]))
      actual.to_s.should == desired.to_s
    end
  end
end
