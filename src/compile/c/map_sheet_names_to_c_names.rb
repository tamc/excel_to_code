class MapSheetNamesToCNames
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,output)
    c_names_assigned = {}
    input.lines do |line|
      excel_worksheet_name = line.split("\t").first
      c_name = excel_worksheet_name.downcase.gsub(/[^a-z0-9]+/,'_')
      c_name = "s"+c_name if c_name[0] !~ /[a-z]/
      c_name = c_name + "2" if c_names_assigned.has_key?(c_name)
      c_name.succ! while c_names_assigned.has_key?(c_name)
      output.puts "#{excel_worksheet_name}\t#{c_name}"
      c_names_assigned[c_name] = excel_worksheet_name
    end
  end
end
