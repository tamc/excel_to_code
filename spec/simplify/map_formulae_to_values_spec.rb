require_relative '../spec_helper'

describe MapFormulaeToValues do
    
  before(:all) do
    @mapper = MapFormulaeToValues.new
  end
  
  checks = test_data('formulae_to_calculated_values.txt').each_line.map.with_index { |line,i| [i,line] }.find_all { |line| line[1].start_with?('[:') }.map { |line| [line[0],*line[1].split("\t")] }
  checks.each do |c|
    it "converts #{c[1].strip} into #{c[2].strip} (line #{c[0]+1} of formulae_to_calculated_values.txt)" do
      desired = eval(c[2])
      actual = @mapper.map(eval(c[1]))
      actual.should == desired
    end
  end
end
