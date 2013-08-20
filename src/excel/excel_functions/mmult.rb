module ExcelFunctions
  
  def mmult(array_a, array_b)
    return array_a if array_a.is_a?(Symbol)
    return array_b if array_b.is_a?(Symbol)
    return :value unless array_a.is_a?(Array)
    return :value unless array_b.is_a?(Array)

    columns = array_a.first.length
    rows = array_b.length
    error = Array.new([rows, columns].max) { Array.new([rows, columns].max, :value) }
    return error unless columns == rows
    return error unless array_a.all? { |a| a.all? { |b| b.is_a?(Numeric) }} 
    return error unless array_b.all? { |a| a.all? { |b| b.is_a?(Numeric) }}
    result = Array.new(array_a.length) { Array.new(array_b.first.length) }
    indices = (0...rows).to_a
    result.map.with_index do |row, i|
      row.map.with_index do |cell, j|
        indices.inject(0) do |sum, n|
          sum = sum + (array_a[i][n] * array_b[n][j])
        end
      end
    end
  end
  
end
