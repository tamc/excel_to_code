# encoding: utf-8
require 'rubypeg'

class Formula < RubyPeg
  
  def root
    formula
  end
  
  def formula
    node :formula do
      optional { space } && one_or_more { expression }
    end
  end
  
  def expression
    string_join || comparison || arithmetic || thing
  end
  
  def thing
    function || array || brackets || any_reference || string || percentage || number || boolean || prefix || error || named_reference
  end
  
  def argument
    expression || null
  end
  
  def function
    node :function do
      terminal(/[A-Z]+/) && ignore { terminal("(") } && space && optional { argument } && any_number_of { (space && ignore { terminal(",") } && space && argument) } && space && ignore { terminal(")") }
    end
  end
  
  def brackets
    node :brackets do
      ignore { terminal("(") } && space && one_or_more { expression } && space && ignore { terminal(")") }
    end
  end
  
  def array
    node :array do
      ignore { terminal("{") } && space && row && any_number_of { (space && ignore { terminal(";") } && space && row) } && space && ignore { terminal("}") }
    end
  end
  
  def row
    node :row do
      basic_type && any_number_of { (space && ignore { terminal(",") } && space && basic_type) }
    end
  end
  
  def basic_type
    string || percentage || number || boolean
  end
  
  def string_join
    node :string_join do
      (arithmetic || thing) && one_or_more { (space && ignore { terminal("&") } && space && (arithmetic || thing)) }
    end
  end
  
  def arithmetic
    node :arithmetic do
      thing && one_or_more { (space && operator && space && thing) }
    end
  end
  
  def comparison
    node :comparison do
      (arithmetic || thing) && space && comparator && space && (arithmetic || thing)
    end
  end
  
  def comparator
    node :comparator do
      terminal(">=") || terminal("<=") || terminal("<>") || terminal(">") || terminal("<") || terminal("=")
    end
  end
  
  def string
    node :string do
      ignore { terminal("\"") } && terminal(/(""|[^"])*/) && ignore { terminal("\"") }
    end
  end
  
  def any_reference
    external_reference || any_internal_reference
  end
  
  def any_internal_reference
    table_reference || local_table_reference || sheet_reference || sheetless_reference
  end
  
  def percentage
    node :percentage do
      terminal(/[-+]?[0-9]+\.?[0-9]*/) && ignore { terminal("%") }
    end
  end
  
  def number
    node :number do
      terminal(/[-+]?[0-9]+\.?[0-9]*([eE][-+]?[0-9]+)?/)
    end
  end
  
  def operator
    node :operator do
      terminal("+") || terminal("-") || terminal("/") || terminal("*") || terminal("^")
    end
  end
  
  def external_reference
    node :external_reference do
      terminal(/\[\d+\]!?/) && any_internal_reference
    end
  end
  
  def table_reference
    node :table_reference do
      table_name && ignore { terminal("[") } && (range_structured_reference || complex_structured_reference || simple_structured_reference) && ignore { terminal("]") }
    end
  end
  
  def local_table_reference
    node :local_table_reference do
      ignore { terminal("[") } && (range_structured_reference || complex_structured_reference || overly_structured_reference || simple_structured_reference) && ignore { terminal("]") }
    end
  end
  
  def table_name
    terminal(/[.\p{Word}_]+/)
  end
  
  def range_structured_reference
    terminal(/\[[^\u005d]*\],\[[^\u005d]*\]:\[[^\u005d]*\]/)
  end
  
  def complex_structured_reference
    terminal(/\[[^\u005d]*\],\[[^\u005d]*\]/)
  end
  
  def overly_structured_reference
    ignore { terminal("[") } && simple_structured_reference && ignore { terminal("]") }
  end
  
  def simple_structured_reference
    terminal(/[^\u005d]*/)
  end
  
  def named_reference
    node :named_reference do
      terminal(/[#\p{word}][\p{word}_.!]+/)
    end
  end
  
  def sheet_reference
    node :sheet_reference do
      (single_quoted_string || terminal(/[\p{word}][\p{word}_.]+/)) && ignore { terminal("!") } && (sheetless_reference || named_reference)
    end
  end
  
  def single_quoted_string
    ignore { terminal("'") } && terminal(/[^']*/) && ignore { terminal("'") }
  end
  
  def sheetless_reference
    column_range || row_range || area || cell
  end
  
  def column_range
    node :column_range do
      column && ignore { terminal(":") } && column
    end
  end
  
  def row_range
    node :row_range do
      row_number && ignore { terminal(":") } && row_number
    end
  end
  
  def area
    node :area do
      reference && ignore { terminal(":") } && reference
    end
  end
  
  def cell
    node :cell do
      reference
    end
  end
  
  def row_number
    terminal(/\$?\d+/)
  end
  
  def column
    terminal(/\$?[A-Za-z]{1,3}/)
  end
  
  def reference
    terminal(/\$?[A-Za-z]{1,3}\$?[0-9]+(?![0-9A-Za-z_])/)
  end
  
  def boolean
    boolean_true || boolean_false
  end
  
  def boolean_true
    node :boolean_true do
      ignore { terminal("TRUE") }
    end
  end
  
  def boolean_false
    node :boolean_false do
      ignore { terminal("FALSE") }
    end
  end
  
  def prefix
    node :prefix do
      terminal(/[-+]/) && thing
    end
  end
  
  def space
    ignore { terminal(/[ \n]*/) }
  end
  
  def null
    node :null do
      followed_by { terminal(",") }
    end
  end
  
  def error
    node :error do
      terminal("#REF!") || terminal("#NAME?") || terminal("#VALUE!") || terminal("#DIV/0!") || terminal("#N/A") || terminal("#NUM!")
    end
  end
  
end
