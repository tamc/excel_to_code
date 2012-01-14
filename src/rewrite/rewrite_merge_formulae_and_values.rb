class RewriteMergeFormulaeAndValues
  def self.rewrite(*args)
    new.rewrite(*args)
  end
  
  def rewrite(formulae,values,output)
    formulae = Hash[formulae.readlines.map { |line| [line[/(.*?)\t/,1],line]}]
    values.lines do |line|
      ref = line[/(.*?)\t/,1]
      output.puts formulae[ref] || line
    end
  end
  
end
  