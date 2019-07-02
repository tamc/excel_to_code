# frozen_string_literal: true

# CombileToGoCommon includes coe shared between CompileToGo and CompileToGoTest
module CompileToGoCommon
  def struct_type
    'spreadsheet'
  end

  def self.rewrite(*args)
    new.rewrite(*args)
  end

  def setup_instance_vars(sheet_names:)
    self.settable ||= ->(_ref) { false }
    self.gettable ||= ->(_ref) { true }
    self.sheet_names = sheet_names
  end

  def getter_method_name(ref)
    v = variable_name(ref).dup
    v[0] = if gettable.call(ref)
             v[0].upcase!
           else
             v[0].downcase!
           end
    v
  end

  def setter_method_name(ref)
    v = variable_name(ref)
    v[0].upcase!
    "Set#{v}"
  end

  def variable_name(ref)
    worksheet = ref.first
    cell = ref.last
    worksheet_name = sheet_names[worksheet.to_s] || worksheet.to_s
    if worksheet_name.empty?
      cell.downcase
    else
      "#{worksheet_name.downcase}#{cell.upcase}"
    end
  end
end

# CompileToGo generates a go version of the passed data structure
class CompileToGo
  include CompileToGoCommon

  attr_accessor :settable
  attr_accessor :gettable
  attr_accessor :sheet_names

  def rewrite(formulae, sheet_names, output)
    setup_instance_vars(sheet_names: sheet_names)

    output.puts struct_definition(formulae: formulae)
    output.puts struct_intializer

    formulae.each do |ref, ast|
      output.puts getter(ref: ref, ast: ast)
      output.puts setter(ref: ref) if settable.call(ref)
    end
  end

  def struct_definition(formulae:)
    ["type #{struct_type} struct {",
     formulae.map do |ref, _|
       "#{variable_name(ref)} cachedValue"
     end,
     '}'].flatten.join("\n")
  end

  def struct_intializer
    <<~ENDGO
       func New() #{struct_type} {
        return #{struct_type}{}
      }
    ENDGO
  end

  def getter(ref:, ast:)
    v = variable_name(ref)
    m = getter_method_name(ref)
    c = code_to_create_value(v, ast)
    <<~ENDGO
      func (s *#{struct_type}) #{m}() (interface{}, error) {
        if !s.#{v}.isCached() {
          #{c}
        }
        return s.#{v}.get()
      }
    ENDGO
  end

  def code_to_create_value(variable_name, ast)
    result = mapper.convert(ast)
    definitions = mapper.get_definitions
    case mapper.result_type
    when :value
      "s.#{variable_name}.set(#{result}, nil)"
    when :error_value
      "s.#{variable_name}.set(nil, #{result})"
    when :function_no_error
      "#{definitions}s.#{variable_name}.set(#{result}, nil)"
    end
  end

  def setter(ref:)
    v = variable_name(ref)
    m = setter_method_name(ref)
    <<~ENDGO
       func (s *#{struct_type}) #{m}(v interface{}) {
        if err, ok := v.(error); ok {
          s.#{v}.set(nil, err)
        } else {
          s.#{v}.set(v, nil)
        }
      }
    ENDGO
  end

  private

  def mapper
    return @mapper if @mapper

    @mapper = MapFormulaeToGo.new
    @mapper.sheet_names = sheet_names
    @mapper.getter_method_name = method(:getter_method_name)
    @mapper
  end
end
