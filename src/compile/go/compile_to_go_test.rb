# frozen_string_literal: true

# Turns the data structure into a series of tests written in golang
class CompileToGoTest
  include CompileToGoCommon

  attr_accessor :settable
  attr_accessor :gettable
  attr_accessor :sheet_names

  def rewrite(formulae, sheet_names, output)
    setup_instance_vars(sheet_names: sheet_names)

    formulae
      .select { |ref, _| gettable.call(ref) }
      .each do |ref, ast|
      if error?(ast: ast)
        output.puts test_for_error(ref: ref, ast: ast)
      else
        output.puts test_for_value(ref: ref, ast: ast)
      end
    end
  end

  def test_for_error(ref:, ast:)
    n = getter_method_name(ref)
    m = MapValuesToGo.new
    <<~ENDGO
      func Test#{n}(t *testing.T) {
        s := New()
        e := #{m.map(ast)}
        a, err := s.#{n}()
        if err != e {
            t.Errorf("#{n} = (%v, %v), want (nil, %v)", a, err, e)
        }
      }
    ENDGO
  end

  def test_for_value(ref:, ast:)
    n = getter_method_name(ref)
    m = MapValuesToGo.new
    <<~ENDGO
      func Test#{n}(t *testing.T) {
        s := New()
        e := #{m.map(ast)}
        a, err := s.#{n}()
        if a != e || err != nil {
            t.Errorf("#{n} = (%v, %v), want (%v, nil)", a, err, e)
        }
      }

    ENDGO
  end

  def error?(ast:)
    ast.first == :error
  end
end
