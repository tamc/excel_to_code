class MapSheetNamesToRubyNames
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,output)
    ruby_names_assigned = {}
    input.lines do |line|
      excel_worksheet_name = line.split("\t").first
      ruby_name = excel_worksheet_name.downcase.gsub(/[^a-z0-9]+/,'_')
      ruby_name = "s"+ruby_name if ruby_name[0] !~ /[a-z]/
      ruby_name = ruby_name + "2" if ruby_names_assigned.has_key?(ruby_name)
      ruby_name.succ! while ruby_names_assigned.has_key?(ruby_name)
      output.puts "#{excel_worksheet_name}\t#{ruby_name}"
      ruby_names_assigned[ruby_name] = excel_worksheet_name
    end
  end
end
