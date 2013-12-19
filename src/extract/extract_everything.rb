require 'ox'

class ExtractEverythingFromWorkbook < ::Ox::Sax

  attr_accessor :table_rids
  attr_accessor :worksheets_dimensions
  attr_accessor :values
  attr_accessor :formulae_simple
  attr_accessor :formulae_array
  attr_accessor :formulae_shared
  attr_accessor :formulae_shared_targets

  def initialize
    # State
    @current_element = []

    # For parsing formulae
    @fp = CachingFormulaParser.instance

    # Outputs
    @table_rids ||= {}
    @worksheets_dimensions ||= {}
    @values ||= {}
    @formulae_simple ||= {}
    @formulae_array ||= {}
    @formulae_shared ||= {}
    @formulae_shared_targets ||= {}
  end

  def extract(sheet_name, input_xml)
    @sheet_name = sheet_name.to_sym
    Ox.sax_parse(self, input_xml, :convert_special => true)
    self
  end
    
  def start_element(name)
    case name
    when :v # :v is value
      @current_element.push(name)
      @value = []
    when :f # :f is formula
      @current_element.push(name)
      @formula = []
      @formula_type = 'simple' # Default is a simple formula, alternatives are shared or array
      @formula_ref = nil # Used to specify range for shared or array formulae
      @si = nil # Used to specify the index for shared formulae
    when :c # :c is cell, wraps values and formulas
      @current_element.push(name)
      @ref = nil
      @value_type = 'n' #Default type is number
    when :dimension, :tablePart
      @current_element.push(name)
    end
  end

  def key
    [@sheet_name, @ref]
  end

  def end_element(name)
    case name
    when :v
      @current_element.pop
      value = @value.join
      ast = case @value_type
      when 'b'; value == "1" ? [:boolean_true] : [:boolean_false]
      when 's'; [:shared_string, value.to_i]
      when 'n'; [:number, value.to_f]
      when 'e'; [:error, value.to_sym]
      when 'str'; [:string, value.gsub(/_x[0-9A-F]{4}_/,'').freeze]
      else
        $stderr.puts "Value of type #{@value_type} not known #{@sheet_name} #{@ref}"
        exit
      end
      @values[key] = @fp.map(ast)
    when :f
      @current_element.pop
      unless @formula.empty?
        formula_text = @formula.join.gsub(/[\r\n]+/,'')
        ast = @fp.parse(formula_text)
        unless ast
          $stderr.puts "Could not parse #{@sheet_name} #{@ref} #{formula_text}"
          exit
        end
      end
      case @formula_type
      when 'simple'
        return if @formula.empty?
        @formulae_simple[key] = ast
      when 'shared'
        @formulae_shared_targets[key] = @si
        if ast
          @formulae_shared[key] = [@formula_ref, @si, ast]
        end
      when 'array'
        @formulae_array[key] = [@formula_ref, ast]
      end
    when :dimension, :tablePart, :c
      @current_element.pop
    end
  end

  def attr(name, value)
    case @current_element.last
    when :dimension
      return unless name == :ref
      @worksheets_dimensions[@sheet_name] = value
    when :tablePart
      return unless name == :'r:id'
      @table_rids[@sheet_name] ||= []
      @table_rids[@sheet_name] << value
    when :c
      case name
      when :r
        @ref = value.to_sym
      when :t
        @value_type = value
      end
    when :f
      case name
      when :t
        @formula_type = value
      when :si
        @si = value
      when :ref
        @formula_ref = value
      end
    end
  end

  def text(text)
    case @current_element.last
    when :v
      @value << text
    when :f
      @formula << text
    end
  end

end
