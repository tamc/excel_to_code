class ReplaceCommonElementsInFormulae
    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,common,output)
    common = common.readlines.map do |a| 
      ref, element = a.split("\t")
      [element.strip,"[:cell, \"#{ref}\"]",ref]
    end.sort
    input.lines do |line|
      common.each do |element,cell,ref|
        line.gsub!(element,cell)
      end
      output.puts line
    end
  end
end
