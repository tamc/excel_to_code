
class CompileToGoTest
  
  attr_accessor :settable
  attr_accessor :gettable
  attr_accessor :sheet_names
  
  def struct_type
    "spreadsheet"
  end

  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(formulae, sheet_names, output)
    self.settable ||= lambda { |ref| false }
    self.gettable ||= lambda { |ref| true }
    self.sheet_names = sheet_names

    m = MapValuesToGo.new

    formulae.each do |ref, ast|
      next unless gettable.call(ref)
      n = getter_method_name(ref)

      if ast.first == :error
        output.puts <<~END
        func Test#{n}(t *testing.T) {
          s := New()
          e := #{m.map(ast)}
          a, err := s.#{n}()
          if err != e {
              t.Errorf("#{n} = (%v, %v), want (nil, %v)", a, err, e)
          }
        }

        END

      else 
        output.puts <<~END
        func Test#{n}(t *testing.T) {
          s := New()
          e := #{m.map(ast)}
          a, err := s.#{n}()
          if a != e || err != nil {
              t.Errorf("#{n} = (%v, %v), want (%v, nil)", a, err, e)
          }
        }

        END
      end
    end
  end

  def getter_method_name(ref)
    v = variable_name(ref)
    if gettable.call(ref)
      v[0] = v[0].upcase!
    else
      v[0] = v[0].downcase!
    end
    v
  end

  def variable_name(ref)
    worksheet = ref.first
    cell = ref.last
    worksheet_name = sheet_names[worksheet.to_s] || worksheet.to_s
    return worksheet_name.length > 0 ? "#{worksheet_name.downcase}#{cell.upcase}" : cell.downcase
  end
  
end
