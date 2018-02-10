
class CompileToGo
  
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

    # The struct
    output.puts "type #{struct_type} struct {"
    formulae.each do |ref, _|
      output.puts "  #{variable_name(ref)} excel.CachedValue"
    end
    output.puts "}"

    # The initializer
    output.puts <<~END

    func New() #{struct_type} {
      return #{struct_type}{}
    }

    END

    formulae.each do |ref, ast|
      v = variable_name(ref)
      output.puts <<~END
        func (s *#{struct_type}) #{getter_method_name(ref)}() (interface{}, error) {
          if !s.#{v}.IsCached() {
            s.#{v}.Set(#{m.map(ast)})
          }
          return s.#{v}.Get()
        }

      END
      if settable.call(ref)
        output.puts <<~END

        func (s *#{struct_type}) #{setter_method_name(ref)}(v interface{}) {
            s.#{v}.Set(v)
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

  def setter_method_name(ref)
    v = variable_name(ref)
    v[0].upcase!
    "Set#{v}"
  end

  def variable_name(ref)
    worksheet = ref.first
    cell = ref.last
    worksheet_name = sheet_names[worksheet.to_s] || worksheet.to_s
    return worksheet_name.length > 0 ? "#{worksheet_name.downcase}#{cell.upcase}" : cell.downcase
  end
  
end
