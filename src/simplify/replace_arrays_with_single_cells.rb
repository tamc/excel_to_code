require_relative '../excel'

class ReplaceArraysWithSingleCellsAst

  attr_accessor :ref
  attr_accessor :need_to_replace

  ERROR = [:error, :"#VALUE!"]

  def map(ast)
    @need_to_replace = false
    return unless ast.is_a?(Array)
    if ast.first == :array
      @need_to_replace = true
      return try_and_convert_array(ast)
    else
      do_map(ast)
      ast
    end
  end

  def do_map(ast)
    return ast unless ast.is_a?(Array)
    case ast.first
    when :arithmetic
      left, op, right = ast[1], ast[2], ast[3]
      if left.first == :array || right.first == :array
        left = try_and_convert_array(left)
        right = try_and_convert_array(right)
        ast.replace([:arithmetic, left, op, right])
      else
        ast[1..-1].each { |a| do_map(a) }
      end
    when :string_join
      strings = ast[1..-1]
      if strings.any? { |s| s.first == :array }
        strings = strings.map { |s| try_and_convert_array(s) }
        ast.replace([:string_join, *strings])
      else
        ast[1..-1].each { |a| do_map(a) }
      end
    else
      ast[1..-1].each { |a| do_map(a) }
    end
  end

  def try_and_convert_array(ast)
    return ast unless ast.first == :array
    #return ast unless ast[1..-1].all? { |c| c.first == :sheet_reference }
    if ast.length == 2
      single_row(ast)
    elsif ast[1].length == 2
      single_column(ast)
    else
      ERROR
    end
  end

  def single_row(ast)
    r = Reference.for(ref.last)
    r.calculate_excel_variables
    column = r.excel_column
    sheet = ref.first

    cells = ast[1][1..-1]
    match = cells.find do |cell|
      s = cell[1]
      c = cell[2][1][/([A-Za-z]{1,3})/,1]
      sheet == s && column == c
    end

    match || ERROR
  end

  def single_column(ast)
    r = Reference.for(ref.last)
    r.calculate_excel_variables
    row = r.excel_row
    sheet = ref.first

    cells = ast[1..-1].map { |row| row.last }
    match = cells.find do |cell|
      s = cell[1]
      r = cell[2][1][/([A-Za-z]{1,3})(\d+)/,2]
      sheet == s && row == r
    end

    match || ERROR
  end
end
