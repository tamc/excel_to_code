require_relative 'map_formulae_to_c'

class MapFormulaeToCFunction < MapFormulaeToC
  
  def cell(reference)
    # FIXME: What a cludge.
    if reference =~ /common\d+/
      "#{reference}"
    else
      reference.to_s.downcase.gsub('$','')
    end
  end
  
  def sheet_reference(sheet,reference)
    "#{sheet_names[sheet]}_#{map(reference).to_s.downcase}"
  end

end
