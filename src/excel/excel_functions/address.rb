module ExcelFunctions

  # ADDRESS(row_num,column_num,abs_num,a1,sheet_text)
  # row_num
  # The row number to use in the cell reference.
  # column_num
  # The column number to use in the cell reference.
  # abs_num
  # Specifies the type of reference to return.
  # 1 or omitted returns the following type of reference: absolute.
  # 2 returns the following type of reference: absolute row; relative column.
  # 3 returns the following type of reference: relative row; absolute column.
  # 4 returns the following type of reference: relative.
  # a1
  # A logical value that specifies the A1 or R1C1 reference style.
  # If a1 is TRUE or omitted, ADDRESS returns an A1-style reference.
  # If a1 is FALSE, ADDRESS returns an R1C1-style reference.
  # sheet_text
  # Text specifying the name of the sheet to be used as the external reference.
  # If sheet_text is omitted, no sheet name is used.  
  def address(row_num, column_num, abs_num = 1, a1 = true, sheet_text = nil)
    raise NotSupportedException.new("address() function R1C1 references not implemented") if a1 == false

    return row_num if row_num.is_a?(Symbol)
    return column_num if column_num.is_a?(Symbol)
    return abs_num if abs_num.is_a?(Symbol)
    return a1 if a1.is_a?(Symbol)
    return sheet_text if sheet_text.is_a?(Symbol)
    
    row_num = number_argument(row_num)
    column_num = number_argument(column_num)
    abs_num = number_argument(abs_num)
    
    return :value if row_num < 1
    return :value if column_num < 1
    return :value if abs_num < 1
    return :value if abs_num > 4
    
    row = row_num.to_i.to_s
    column = Reference.column_letters_for_column_number[column_num.to_i].to_s
    
    ref = case abs_num
    when 1; "$"+column+"$"+row
    when 2; "" +column+"$" +row
    when 3; "$"+column+"" +row
    when 4;  ""+column+"" +row
    end
    
    if sheet_text
      ref = "'#{sheet_text}'!"+ref
    end
    
    ref

  end
  
end
