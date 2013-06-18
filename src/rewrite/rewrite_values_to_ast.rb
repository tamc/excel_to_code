require_relative '../excel/formula_peg'

class RewriteValuesToAst
  
  def self.rewrite(input,output)
    self.new.rewrite(input,output)
  end
  
  # input should be in the form: 'thing\tthing\tformula\n' where the last field is always a forumla
  # output will be in the form 'thing\tthing\tast\n'
  def rewrite(input,output)
    input.each_line do |line|
      line =~ /^(.*?)\t(.*?)\t(.*)\n/
      ref, type, value = $1, $2, $3
      ast = case type
      when 'b'; value == "1" ? [:boolean_true] : [:boolean_false]
      when 's'; [:shared_string,value]
      when 'n'; [:number,value]
      when 'e'; [:error,value]
      when 'str'; [:string,value]
      else
        $stderr.puts "Type #{type} not known in #{line}"
        [:parse_error,line.inspect]
      end
      output.puts "#{ref}\t#{ast}"
    end
  end
end
