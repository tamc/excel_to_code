module ExcelFunctions

  def apply_to_range(a,b)
    a = Array.new(b.length,Array.new(b.first.length,a)) unless a.is_a?(Array)
    b = Array.new(a.length,Array.new(a.first.length,b)) unless b.is_a?(Array)
    
    return :value unless b.length == a.length
    return :value unless b.first.length == a.first.length
    
    a.map.with_index do |row,i|
      row.map.with_index do |cell,j|
        yield cell, b[i][j]
      end
    end
  end
  
end
