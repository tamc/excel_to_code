
class RemoveCells
  
  attr_accessor :cells_to_keep
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(formulae)
    formulae.delete_if do |ref, ast|
      delete_ref?(ref)
    end
    formulae
  end

  def delete_ref?(ref)
    sheet = ref.first
    cell = ref.last
    cells_to_keep_in_sheet = cells_to_keep[sheet]
    return true unless cells_to_keep_in_sheet
    return false if cells_to_keep_in_sheet.has_key?(cell)
    true
  end
end
