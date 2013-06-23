
class RemoveCells
  
  attr_accessor :cells_to_keep
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,output)
    input.each_line do |line|
      ref = line[/^(.*?)\t/,1]
      if cells_to_keep.has_key?(ref)
        output.puts line
      end
    end
  end
end
