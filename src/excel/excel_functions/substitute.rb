module ExcelFunctions
  
  def substitute(text, old_text, new_text, instance_num = :any)
    # Check for errors
    return text if text.is_a?(Symbol)
    return old_text if old_text.is_a?(Symbol)
    return new_text if new_text.is_a?(Symbol)
    # Nils get turned into blanks
    text ||= ""
    new_text ||= ""
    old_text ||= ""
    return text if old_text == ""
    # Booleans get made into text, but need to be TRUE not true
    text = text.to_s.upcase if text.is_a?(TrueClass) || text.is_a?(FalseClass)
    old_text = old_text.to_s.upcase if old_text.is_a?(TrueClass) || old_text.is_a?(FalseClass)
    new_text = new_text.to_s.upcase if new_text.is_a?(TrueClass) || new_text.is_a?(FalseClass)
    # Other items get turned into text
    text = text.to_s
    old_text = old_text.to_s
    new_text = new_text.to_s
    # Now split based on whether instance_num is passed
    if instance_num == :any
      # The easy case
      text.gsub(old_text, new_text)
    else
      # The harder case
      return instance_num if instance_num.is_a?(Symbol)
      return :value unless instance_num.to_i > 0
      instances_found = 0
      text.gsub(old_text) do |match|
        instances_found += 1
        if instances_found == instance_num
          new_text
        else
          old_text
        end
      end
    end
  end
  
end
