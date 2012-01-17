require_relative '../excel/formula_peg'

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
        output.puts ast.to_ast[1].to_s
      else
        $stderr.puts "Formula not parsed on line #{i+1}: #{line}"
        output.puts "[:parse_error,#{line.inspect}]"
      end
    end
  end
end
