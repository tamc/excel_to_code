require_relative '../excel'

class ReplaceRangesWithArrayLiteralsAst
  def initialize
    @cache = {}
  end

  def map(ast)
    r = do_map(ast)
    ast.replace(r) unless r.object_id == ast.object_id
    ast
  end

  def do_map(ast)
    return ast unless ast.is_a?(Array)
    case ast[0]
    when :sheet_reference; return sheet_reference(ast)
    when :area; return area(ast)
    when :function;
      if ast[1] == :SUMIF
        return sumif(ast)
      else
        map_args(ast)
      end
    else
      map_args(ast)
    end
    ast
  end

  def map_args(ast)
    ast.each.with_index do |a,i|
      next unless a.is_a?(Array)
      case a[0]
      when :sheet_reference; ast[i] = sheet_reference(a)
      when :area; ast[i] = area(a)
      when :function
        if ast[1] == :sumif
          ast[i] = sumif(a)
        else
          do_map(a) 
        end
      else 
        do_map(a)
      end
    end
  end

  # ARGH: SUMIF(A1:A10, 10, B5:B6) is interpreted by Excel as SUMIF(A1:A10, 10, B5:B15)
  def sumif(ast)
    # If only two arguments to SUMIF, won't be a problem
    return map_args(ast) unless ast.length == 5
    check_range, criteria, sum_range = ast[2], ast[3], ast[4]
    return map_args(ast) unless [:area, :sheet_reference, :cell].include?(check_range.first)
    return map_args(ast) unless [:area, :sheet_reference, :cell].include?(sum_range.first)
    check_area = area_for(check_range).unfix
    sum_area = area_for(sum_range).unfix

    check_area.calculate_excel_variables
    sum_area.calculate_excel_variables

    return map_args(ast) if check_area.height  == sum_area.height && check_area.width == sum_area.width

    new_sum_area = [:area, sum_area.excel_start.to_sym, sum_area.excel_start.offset(check_area.height, check_area.width).to_sym]

    if sum_range.first == :sheet_reference
      ast[4][2] = new_sum_area
    else
      ast[4] = new_sum_area
    end
    map_args(ast)
  end

  def area_for(ast)
    case ast.first
    when :cell then Area.for(ast[1])
    when :area then Area.for("#{ast[1]}:#{ast[2]}")
    when :sheet_reference then area_for(ast[2])
    end
  end

  # Of the form [:sheet_reference, sheet, reference]
  def sheet_reference(ast)
    @cache[ast] || calculate_expansion_for(ast)
  end

  def calculate_expansion_for(ast)
    sheet = ast[1]
    reference = ast[2]
    return ast unless reference.first == :area
    area = Area.for("#{reference[1]}:#{reference[2]}")
    a = area.to_array_literal(sheet)

    # Don't convert single cell ranges
    result = if a.size == 2 && a[1].size == 2
               a[1][1]
             else
               a
             end
    @cache[ast.dup] = result
  end

  # Of the form [:area, start, finish]
  def area(ast)
    start = ast[1]
    finish = ast[2]
    area = Area.for("#{start}:#{finish}")
    a = area.to_array_literal

    # Don't convert single cell ranges
    if a.size == 2 && a[1].size == 2
      a[1][1]
    else
      a
    end
  end

end

class ReplaceRangesWithArrayLiterals

  def self.replace(*args)
    self.new.replace(*args)
  end

  def replace(input,output)
    rewriter = ReplaceRangesWithArrayLiteralsAst.new

    input.each_line do |line|
      # Looks to match shared string lines
      if line =~ /\[:area/
        content = line.split("\t")
        ast = eval(content.pop)
        output.puts "#{content.join("\t")}\t#{rewriter.map(ast).inspect}"
      else
        output.puts line
      end
    end
  end

end
