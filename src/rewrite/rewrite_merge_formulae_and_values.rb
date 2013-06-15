class RewriteMergeFormulaeAndValues
  def self.rewrite(*args)
    new.rewrite(*args)
  end
  
  attr_accessor :references_to_add_if_they_are_not_already_present
  
  def rewrite(values,shared_formulae,array_formula,simple_formulae,output)
    @references_to_add_if_they_are_not_already_present ||= []
    
    shared_formulae = Hash[shared_formulae.readlines.map { |line| [line[/(.*?)\t/,1],line]}]
    array_formula = Hash[array_formula.readlines.map { |line| [line[/(.*?)\t/,1],line]}]
    simple_formulae = Hash[simple_formulae.readlines.map { |line| [line[/(.*?)\t/,1],line]}]
        
    values.each_line do |line|
      ref = line[/(.*?)\t/,1]
      @references_to_add_if_they_are_not_already_present.delete(ref)
      output.puts simple_formulae[ref] || array_formula[ref] || shared_formulae[ref] || line
    end
    @references_to_add_if_they_are_not_already_present.each do |r|
      output.puts "#{r}\t[:blank]"
    end
  end
  
end
  
