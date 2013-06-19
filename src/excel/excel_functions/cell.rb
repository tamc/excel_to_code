module ExcelFunctions
  
  def cell(info_type, reference = nil)
    return info_type if info_type.is_a?(Symbol)
    return :value unless info_type.is_a?(String)
    case info_type.downcase
    when 'filename'
      original_excel_filename
    else
      :value
    end
  end
  
end
