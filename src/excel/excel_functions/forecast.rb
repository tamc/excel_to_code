module ExcelFunctions
  
  def forecast(required_x, known_y, known_x)
    required_x = number_argument(required_x)
    return required_x if required_x.is_a?(Symbol)
    fit = linest(known_y, known_x)
    return fit if fit.is_a?(Symbol)
    fit = fit.first
    intercept = fit.first
    slope = fit.last
    return intercept + (slope*required_x)
  end

  def linest(known_y, known_x)
    return :na unless known_y.is_a?(Array)
    return :na unless known_x.is_a?(Array)
    known_y = known_y.flatten
    known_x = known_x.flatten
    known_y.each { |y| return y if y.is_a?(Symbol) }
    known_x.each { |x| return x if x.is_a?(Symbol) }
    return :na unless known_x.length == known_y.length
    return :na if known_y.empty?
    return :na if known_x.empty?

    0.upto(known_x.length-1).each do |i|
      known_x[i] = nil unless known_y[i].is_a?(Numeric)
      known_y[i] = nil unless known_x[i].is_a?(Numeric)
    end

    known_x.compact!
    known_y.compact!

    mean_y = known_y.inject(0.0) { |m,i| m + i.to_f }   / known_y.size.to_f
    mean_x = known_x.inject(0.0) { |m,i| m + i.to_f }   / known_x.size.to_f

    b_denominator = known_x.inject(0.0) { |s,x| x.is_a?(Numeric) ? (s + (x-mean_x)**2 ) : s  }
    return :div0 if b_denominator == 0
    b_numerator = 0.0
    known_x.each.with_index do |x,i|
      y = known_y[i]
      next unless x.is_a?(Numeric)
      next unless y.is_a?(Numeric)
      b_numerator = b_numerator + ((x-mean_x)*(y-mean_y))
    end

    b = b_numerator / b_denominator
    a = mean_y - (b * mean_x)

    return [[a, b]]
  end
  
end
