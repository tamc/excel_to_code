# frozen_string_literal: true

# Generates a go version of the passed data structure
class CompileToGo
  attr_accessor :settable
  attr_accessor :gettable
  attr_accessor :sheet_names

  def struct_type
    'spreadsheet'
  end

  def self.rewrite(*args)
    new.rewrite(*args)
  end

  def rewrite(formulae, sheet_names, output)
    setup_instance_vars(sheet_names: sheet_names)

    output.puts struct_definition(formulae: formulae)
    output.puts struct_intializer

    formulae.each do |ref, ast|
      output.puts getter(ref: ref, ast: ast)
      output.puts setter(ref: ref) if settable.call(ref)
    end
  end

  def setup_instance_vars(sheet_names:)
    self.settable ||= ->(_ref) { false }
    self.gettable ||= ->(_ref) { true }
    self.sheet_names = sheet_names
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
    value = MapValuesToGo.new.map(ast)
    <<~ENDGO
      func (s *#{struct_type}) #{m}() (interface{}, error) {
        if !s.#{v}.isCached() {
          s.#{v}.set(#{value})
        }
        return s.#{v}.get()
      }
    ENDGO
  end

  def setter(ref:)
    v = variable_name(ref)
    m = setter_method_name(ref)
    <<~ENDGO
       func (s *#{struct_type}) #{m}(v interface{}) {
          s.#{v}.set(v)
      }
    ENDGO
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
