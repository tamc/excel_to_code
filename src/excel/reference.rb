class Reference
  
  @@column_number_for_column ||= Hash.new do |hash,letters|
    number = letters.downcase.each_byte.to_a.reverse.each.with_index.inject(0) do |memo,byte_with_index,c|
      memo + ((byte_with_index.first - 96) * (26**byte_with_index.last))
    end
    hash[letters] = number
    @@column_letters_for_column_number[number] = letters.upcase
    number
  end
  
  @@column_letters_for_column_number ||= Hash.new do |hash,number|
    letters = (number-1).to_i.to_s(26)
    letters = (letters[0...-1].tr('1-9a-z','abcdefghijklmnopqrstuvwxyz') + letters[-1,1].tr('0-9a-z','abcdefghijklmnopqrstuvwxyz')).gsub('a0','z').gsub(/([b-z])0/) { $1.tr('b-z','a-y')+"z" }
    letters.upcase!
    hash[number] = letters
    @@column_number_for_column[letters] = number
    letters
  end
  
  def initialize(text)
    text =~ /(\$)?([A-Za-z]{1,3})(\$)?([0-9]+)/
    @fc, @c, @fr, @r = $1, $2, $3, $4
  end
    
  def offset(rows,columns)
    dup.offset!(rows,columns)
  end
  
  def offset!(rows,columns)
    @r = @r.to_i + rows
    @c = @@column_letters_for_column_number[@@column_number_for_column[@c] + columns]
    self
  end
  
  def to_s
    "#{@fc && '$'}#{@c}#{@fr && '$'}#{@r}"
  end

end