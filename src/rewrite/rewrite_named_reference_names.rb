class RewriteNamedReferenceNames

  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  # Expects named reference in the form:
  # sheet_name_or_blank_for_global\tnamed_reference_name\treference\n
  # Outputs named references in the form:
  # name\treference\n
  # where name incorporates the sheet name and has been rewritten in a way
  # that works as a c function name and (hopefully) won't clash with any
  # existing names 
  # FIXME: but could still clash with function names and methods in the ruby shim
  def rewrite(named_references, worksheet_names, output)
    worksheet_names = Hash[worksheet_names.readlines.map { |line| line.strip.split("\t")}]
    c_names_assigned = worksheet_names.invert

    named_references.each_line do |line|
      sheet, name, reference = line.split("\t")
      sheet = worksheet_names[sheet]
      if sheet
        c_name = "#{sheet}_#{name}"
      else
        c_name = name
        c_name = "n"+c_name if c_name[0] !~ /[a-zA-Z]/
      end
      c_name = c_name.downcase.gsub(/[^a-z0-9]+/,'_')
      c_name = c_name + "2" if c_names_assigned.has_key?(c_name)
      c_name.succ! while c_names_assigned.has_key?(c_name)
      output.puts "#{c_name}\t#{reference}"
      c_names_assigned[c_name] = c_name
    end
  end
end
