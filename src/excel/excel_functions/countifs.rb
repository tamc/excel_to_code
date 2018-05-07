module ExcelFunctions
  
  def countifs(*criteria)
    return 0 if criteria.empty?
    c = criteria[0].is_a?(Array) ? criteria[0] : [criteria[0]]
    count = c.map { 1 }
    return sumifs(count, *criteria)
  end
  
end
