class RewriteRelationshipIdToFilename
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,relationships_file,output)
    relationships_file.rewind
    relationships_file.rewind
    relationships = Hash[relationships_file.readlines.map { |line| line.split("\t")}]
    input.lines do |line|
      parts = line.split("\t")
      rid = parts.pop.strip
      if relationships.has_key?(rid)
        output.puts "#{parts.join("\t")}#{parts.size > 0 ? "\t" : ""}#{relationships[rid].strip}"
      else
        $stderr.puts "Warning, #{rid.inspect} not found in relationships file #{relationships.inspect}"
        outputs.puts "Warning, #{rid.inspect} not found in relationships file #{relationships.inspect}"
      end
    end
  end
end
