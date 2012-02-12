class RewriteMergeFormulaeAndValues
  def self.rewrite(*args)
    new.rewrite(*args)
  end
  
  def rewrite(values,shared_formulae,array_formula,simple_formulae,output)
    shared_formulae = Hash[shared_formulae.readlines.map { |line| [line[/(.*?)\t/,1],line]}]
    array_formula = Hash[array_formula.readlines.map { |line| [line[/(.*?)\t/,1],line]}]
    simple_formulae = Hash[simple_formulae.readlines.map { |line| [line[/(.*?)\t/,1],line]}]

    values.lines do |line|
      ref = line[/(.*?)\t/,1]
      output.puts simple_formulae[ref] || array_formula[ref] || shared_formulae[ref] || line
    end
  end
  
end
  