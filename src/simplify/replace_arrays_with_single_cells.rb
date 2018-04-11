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
      new_ast = try_and_convert_array(ast)
      return ERROR if new_ast.first == :array
      return new_ast
    # Special case, only change if at the top level
    elsif ast[0] == :function && ast[1] == :CHOOSE && check_choose(ast)
      # Replacement made in check
    # Special case, only change if at the top level
    elsif ast[0] == :function && ast[1] == :IF && check_if(ast)
      # Replacement made in check
    else
      do_map(ast)
    end
    ast
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
        map_if_required(ast)
      end
    when :prefix
      op, left = ast[1], ast[2]
      if left.first == :array
        left = try_and_convert_array(left)
        ast.replace([:prefix, op, left])
      else
        map_if_required(ast)
      end
    when :string_join
      strings = ast[1..-1]
      if strings.any? { |s| s.first == :array }
        strings = strings.map { |s| try_and_convert_array(s) }
        ast.replace([:string_join, *strings])
      else
        map_if_required(strings)
      end
    when :function
      if ast[1] == :SUMIF && ast[3].first == :array
        ast[3] = try_and_convert_array(ast[3])
      elsif ast[1] == :SUMIFS && check_sumifs(ast)
        # Replacement madein check_sumif function
      elsif ast[1] == :MATCH && check_match(ast)
        # Replacement made in check_match function
      elsif ast[1] == :INDIRECT && check_indirect(ast)
        # Replacement made in check function
      elsif ast[0] == :function && ast[1] == :INDEX && check_index(ast)
        # Replacement made in check
      else
        map_if_required(ast)
      end
    else
      map_if_required(ast)
    end
  end

  # This tries to speed up the process by not recursively
  # checking arguments that we know won't be interesting
  def map_if_required(arguments)
    return arguments unless arguments.is_a?(Array)
    arguments.each do |a|
      next unless a.is_a?(Array)
      case a.first
      when :error, :null, :space, :prefix, :boolean_true, :boolean_false, :number, :string
        next
      when :sheet_reference, :table_reference, :local_table_reference
        next
      else
        do_map(a)
      end
    end
  end

  def check_choose(ast)
    replacement_made = false
    i = 2
    while i < ast.length
      if ast[i].first == :array
        replacement_made = true
        ast[i] = try_and_convert_array(ast[i])
      end
      i +=1
    end
    replacement_made
  end

  def check_if(ast)
    replacement_made = false
    if ast[3] && ast[3].first == :array
      replacement_made = true
      ast[3] = try_and_convert_array(ast[3])
    end
    if ast[4] && ast[4].first == :array
      replacement_made = true
      ast[4] = try_and_convert_array(ast[4])
    end
    replacement_made
  end

  def check_index(ast)
    replacement_made = false
    if ast[3] && ast[3].first == :array
      replacement_made = true
      ast[3] = try_and_convert_array(ast[3])
    end
    if ast[4] && ast[4].first == :array
      replacement_made = true
      ast[4] = try_and_convert_array(ast[4])
    end
    replacement_made
  end

  def check_match(ast)
    return false unless ast[2].first == :array
    ast[2] = try_and_convert_array(ast[2])
    ast[3..-1].each { |a| do_map(a) }
    true
  end

  def check_indirect(ast)
    return false unless ast[2].first == :array
    ast[2] = try_and_convert_array(ast[2])
    true
  end

  def check_sumifs(ast)
    replacement_made = false
    i = 4
    while i < ast.length
      if ast[i].first == :array
        replacement_made = true
        ast[i] = try_and_convert_array(ast[i])
      end
      i +=2
    end
    replacement_made
  end

  def try_and_convert_array(ast)
    return ast unless ast.first == :array
    return ast unless all_references?(ast)
    #return ast unless ast[1..-1].all? { |c| c.first == :sheet_reference }
    if ast.length == 2
      single_row(ast)
    elsif ast[1].length == 2
      single_column(ast)
    else
      ast
    end
  end

  def all_references?(ast)
    ast[1..-1].all? do |row|
      row[1..-1].all? do |cell|
        cell.original.first == :sheet_reference
      end
    end
  end

  def single_row(ast)
    r = Reference.for(ref.last)
    r.calculate_excel_variables
    column = r.excel_column
    sheet = ref.first

    cells = ast[1][1..-1]
    match = cells.find do |cell|
      s = cell.original[1]
      c = cell.original[2][1][/([A-Za-z]{1,3})/,1]
      sheet == s && column == c
    end

    match || ast
  end

  def single_column(ast)
    r = Reference.for(ref.last)
    r.calculate_excel_variables
    row = r.excel_row
    sheet = ref.first

    cells = ast[1..-1].map { |row| row.last }
    match = cells.find do |cell|
      s = cell.original[1]
      r = cell.original[2][1][/([A-Za-z]{1,3})(\d+)/,2]
      sheet == s && row == r
    end

    match || ast
  end
end
