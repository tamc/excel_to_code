class RewriteRelationshipIdToFilename
  
  def self.rewrite(worksheet_names,relationships,output)
    self.new.rewrite(worksheet_names,relationships,output)
  end
  
  def rewrite(input,relationships,output)
    relationships = Hash[relationships.readlines.map { |line| line.split("\t")}]
    input.lines do |line|
      parts = line.split("\t")
      rid = parts.pop.strip
      output.puts "#{parts.join("\t")}#{parts.size > 0 ? "\t" : ""}#{relationships[rid].strip}"
    end
  end
end
