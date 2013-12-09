# FIXME Make this a subclass of symbol
class Reference < String
  
  # This is so that we only have one instance for a given reference
  @@references_for_text ||= Hash.new do |hash,text|
    hash[text] = reference = Reference.new(text)
    reference
  end
    
  # This is so that we only have one instance of a given reference specified by its variables
  def Reference.for(text)
     @@references_for_text[text.to_s]
  end
    
  # This caches the calculation for turning column letters (e.g., AAB) into column numbers (e.g., 127)
  @@column_number_for_column ||= Hash.new do |hash,letters|
    number = letters.downcase.each_byte.to_a.reverse.each.with_index.inject(0) do |memo,byte_with_index,c|
      memo + ((byte_with_index.first - 96) * (26**byte_with_index.last))
    end
    hash[letters] = number
    @@column_letters_for_column_number[number] = letters.upcase
    number
  end
  
  # This caches the calculation for turning column numbers (e.g., 127) into column letters (e.g., AAB)
  @@column_letters_for_column_number ||= Hash.new do |hash,number|
    letters = (number-1).to_i.to_s(26)
    letters = (letters[0...-1].tr('1-9a-z','abcdefghijklmnopqrstuvwxyz') + letters[-1,1].tr('0-9a-z','abcdefghijklmnopqrstuvwxyz')).gsub('a0','z').gsub(/([b-z])0/) { $1.tr('b-z','a-y')+"z" }
    letters.upcase!
    hash[number] = letters
    @@column_number_for_column[letters] = number
    letters
  end
  
  attr_reader :excel_row_number, :excel_column_number, :excel_column, :excel_row
  
  def calculate_excel_variables
    return if @excel_variables_calculated
    self =~ /(\$)?([A-Za-z]{1,3})(\$)?([0-9]+)/
    @excel_fixed_column, @excel_column, @excel_fixed_row, @excel_row = $1, $2, $3, $4
    @excel_row_number = @excel_row.to_i
    @excel_column_number = @@column_number_for_column[@excel_column]
    @excel_variables_calculated = true
  end
    
  def offset(rows,columns)
    calculate_excel_variables
    new_column = @excel_fixed_column ? @excel_column :  @@column_letters_for_column_number[@excel_column_number + columns]
    new_row = @excel_fixed_row ? @excel_row : @excel_row_number + rows
    [@excel_fixed_column,new_column,@excel_fixed_row,new_row].join.to_sym
  end
  
  def unfix
    gsub("$","")
  end
  
end
