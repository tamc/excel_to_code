module ExcelFunctions
  
  def sumproduct(*args)
    error = args.find { |a| a.is_a?(Symbol) }
    return error if error
    return :value if args.any? { |a| a == nil }
    args = args.map { |a| a.is_a?(Array) ? a : [[a]] }
    first = args.shift
    accumulator = 0
    first.each_with_index do |row,row_number|
      row.each_with_index do |cell,column_number|
        next unless cell.is_a?(Numeric)
        product = cell
        args.each do |area|
          return :value unless area.length > row_number
          r = area[row_number]
          return :value unless r.length > column_number
          c = r[column_number]
          if c.is_a?(Numeric)
            product = product * c
          else
            product = product * 0
            break
          end
        end
        accumulator += product
      end
    end
    accumulator
  end
  
end
