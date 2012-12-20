require_relative '../../util/not_supported_exception'

class MapForumlaeToLinkedHTML

  def map(ast)
    return ast unless ast.is_a?(Array) # i.e., a value such as 1
    if ast.first.is_a?(Symbol) # i.e., [:operator, '+']
      operator = ast[0]
      arguments = ast[1..-1]
      if respond_to?(operator)
        send(operator,*arguments)
      else
        default(operator, arguments)
      end
    else # i.e., a stream of arguments [[:number, '1'], [:operator, '+'], [:number, '2']]
      ast.map do |a|
        map(a)
      end
    end
  end
  
  def default(operator, arguments)
    "[#{operator}, #{map(arguments).join(", ")}]"
  end

  def function(name, *arguments)
    "#{name.upcase}(#{map(arguments).join(', ')})"
  end

  def brackets(*arguments)
    "(#{map(arguments).join('')})"
  end

  def string_join(left, right)
    "#{map(left)}&#{map(right)}"
  end

  def arithmetic(*arguments)
    map(arguments).join('')
  end

  def comparison(*arguments)
    map(arguments).join('')
  end

  def string(s)
    s.inspect
  end

  def percentage(p)
    "#{p*100}%"
  end

  def number(n)
    n
  end

  def operator(op)
    op
  end

  def sheet_reference(sheet,reference)
    @sheet = sheet+".html"
    s = map(reference)
    @sheet = nil
    s
  end
  
  def area(start,finish)
    cell(start)+":"+cell(finish)
  end

  def cell(ref)
    "<a href=\"#{@sheet}##{ref.gsub('$','')}\">#{ref}</a>"
  end

  def boolean_false()
    'FALSE'
  end

  def boolean_true()
    'TRUE'
  end

  def prefix(op, *arguments)
    op + map(arguments).join('')
  end

  def null()
    ', '
  end



end


class CompileToHTML
  
  attr_accessor :dimensions
  attr_accessor :formulae
  attr_accessor :values

  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(sheet_name,o)
    cells = Area.for(dimensions[sheet_name]).to_array_literal

    # Put in the header and preamble
    o.puts <<-END
    <html>
      <link href='application.css' rel='stylesheet' type='text/css' />
      <script type='text/javascript' src='jquery.min.js'></script>
      <script type='text/javascript' src='application.js'></script>
      <body>

      <div id='formulabar'>
        [<a id='workbook' href=''>2050Model.xlsx</a>]'#{sheet_name}'!<span id='selectedcell'></span>=<span id='selectedformula'></span> Value = <span id='selectedvalue'></span>
      </div>
    END

    # Put in the worksheet
    o.puts "<table class='cells'>"
    cells.shift # :array

    # Put in the header row
    o.puts "<tr>"
    cells.first.each do |cell|
      if cell.is_a?(Array)
        o.puts "<th>#{cell.last[/[a-zA-Z]+/]}</th>"
      else
        o.puts "<th></th>"
      end
    end
    o.puts "</tr>"

    # Put in the actual content
    cells.each do |row|
      o.puts "<tr>"
      # Put in the row number
      o.puts "<th>#{row[1].last[/[0-9]+/]}</th>"
      row.shift # :row
      row.each do |cell|
        ref = cell.last
        o.puts "<td id='#{ref}' data-formula='#{formula(sheet_name,ref)}' data-value='#{value(sheet_name, ref)}'>#{formatted_value(sheet_name, ref)}</td>"
      end
    end
    o.puts "</table>"
    o.puts "<p>Generated on #{Time.now} by <a href='http://github.com/tamc/excel_to_code'>excel_to_code</a></p>"

    o.puts "<div id='jumpbar'>Worksheets: "
    dimensions.each do |name, dimensions|
      o.puts "<a href='#{name}.html' class='#{name == sheet_name && "current"}'>#{name}</a>"
    end
    o.puts "</div>"

    # Put in the closing tags
    o.puts "</body>"
    o.puts "</html>"
  end
  
  def formatted_value(sheet, cell)
    v = values[sheet][cell]
    return nil unless v
    case v.first
    when :number
      v.last.to_f.round
    else
      v.last
    end
  end

  def value(sheet, cell)
    v = values[sheet][cell]
    return nil unless v
    v.last
  end

  def formula(sheet, cell)
    f = formulae[sheet][cell]
    return nil unless f
    ast_to_html(f)
  end

  def ast_to_html(ast)
    @mapper ||= MapForumlaeToLinkedHTML.new
    @mapper.map(ast)
  end


  def worksheet_dimensions=(worksheet_dimensions)
    @dimensions =  Hash[worksheet_dimensions.readlines.map do |line| 
      worksheet_name, area = line.split("\t")
      [worksheet_name,area]
    end]
    @mapper = nil
  end
end
