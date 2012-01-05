class NamedReferences
  
  attr_accessor :named_references
  
  def initialize(refs)
    @named_references = {}
    refs.each do |line|
      sheet, name, reference = line.split("\t")
      @named_references[sheet] ||= {}
      @named_references[sheet][name] = eval(reference)
    end
  end
  
  def reference_for(sheet,named_reference)
    if @named_references.has_key?(sheet)
      @named_references[sheet][named_reference] || @named_references[""][named_reference]
    else
      @named_references[""][named_reference]
    end
  end
  
end

class ReplaceNamedReferencesAst
  
  attr_accessor :named_references, :default_sheet_name
  
  def initialize(named_references, default_sheet_name)
    @named_references, @default_sheet_name = named_references, default_sheet_name
  end
  
  def map(ast)
    if ast.is_a?(Array)
      operator = ast.shift
      if respond_to?(operator)
        send(operator,*ast)
      else
        [operator,*ast.map {|a| map(a) }]
      end
    else
      return ast
    end
  end
  
  def quoted_sheet_reference(sheet,reference)
    if reference.first == :named_reference
      named_references.reference_for(sheet,reference.last)
    else
      [:quoted_sheet_reference,sheet,reference]
    end
  end
  
  def sheet_reference(sheet,reference)
    if reference.first == :named_reference
      named_references.reference_for(sheet,reference.last)
    else
      [:sheet_reference,sheet,reference]
    end    
  end
  
  def named_reference(name)
    named_references.reference_for(default_sheet_name,name)
  end
  
end
  

class ReplaceNamedReferences
  
  def self.replace(values,sheet_name,named_references,output)
    self.new.replace(values,sheet_name,named_references,output)
  end
  
  # Rewrites ast with named references
  def replace(values,sheet_name,named_references,output)
    named_references = NamedReferences.new(named_references.readlines)
    rewriter = ReplaceNamedReferencesAst.new(named_references,sheet_name)
    values.lines do |line|
      # Looks to match shared string lines
      if line =~ /\[:named_reference/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
  end
end
