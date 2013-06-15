class RewriteRelationshipIdToFilename
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input, relationships_file, output)
    relationships_file.rewind
    relationships = Hash[relationships_file.readlines.map { |line| line.split("\t")}]
    input.each_line do |line|
      parts = line.split("\t")
      rid = parts.pop.strip
      if relationships.has_key?(rid)
        parts.push relationships[rid].strip
        output.puts parts.join("\t")
      else
        $stderr.puts "Warning, #{rid.inspect} not found in relationships file #{relationships.inspect}"
        output.puts "Warning, #{rid.inspect} not found in relationships file #{relationships.inspect}"
      end
    end
  end
end
