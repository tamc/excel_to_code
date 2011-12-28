class RewriteValuesToIncludeSharedStrings
  
  def self.rewrite(values,shared_strings,output)
    self.new.rewrite(values,shared_strings,output)
  end
  
  # Rewrites values with type 's' (i.e., shared strings) to have type 'str' (i.e., not shared strings)
  def rewrite(values,shared_strings,output)
    shared_strings = shared_strings.readlines
    values.lines do |line|
      # Looks to match shared string lines of the form A1 s 0
      if line =~ /^(.*?)\ts\t(.*)$/
        output.puts "#{$1}\tstr\t#{shared_strings[$3.to_i]}"
      else
        output.puts line
      end
    end
  end
end
