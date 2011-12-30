require_relative 'formula_peg'

class RewriteFormulaeToAst
  
  def self.rewrite(input,output)
    self.new.rewrite(input,output)
  end
  
  # input should be in the form: 'thing\tthing\tformula\n' where the last field is always a forumla
  # output will be in the form 'thing\tthing\tast\n'
  def rewrite(input,output)
    input.lines.with_index do |line,i|
      line =~ /^(.*\t)(.*?)$/
      output.write $1
      ast =  Formula.parse($2)
      if ast
        output.puts  Formula.parse($2).to_ast.to_s
      else
        $stderr.puts "Formula not parsed on line #{i+1}: #{line}"
        output.puts "[:formula, [:string,#{line.inspect}]"
      end
    end
  end
end
