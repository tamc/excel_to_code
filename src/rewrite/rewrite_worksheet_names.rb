class RewriteWorksheetNames
  
  def self.rewrite(worksheet_names,relationships,output)
    self.new.rewrite(worksheet_names,relationships,output)
  end
  
  # Expects worksheet names in the form:
  # name\trelationship_id\n
  # Expects relationships in the form:
  # relationship_id\tfilename\n
  # Outputs worksheet names in the form:
  # name\tfilename\n
  def rewrite(worksheet_names,relationships,output)
    relationships = Hash[relationships.readlines.map { |line| line.split("\t")}]
    worksheet_names.lines do |line|
      rid, name = line.split("\t")
      output.puts "#{name.strip}\t#{relationships[rid].strip}"
    end
  end
end