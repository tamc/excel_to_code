module ExcelFunctions
end

# Support functions
require_relative 'excel_functions/number_argument'

# Constants
require_relative 'excel_functions/pi'

# Comparators
require_relative 'excel_functions/excel_equal'
require_relative 'excel_functions/less_than'
require_relative 'excel_functions/more_than'
require_relative 'excel_functions/less_than_or_equal'
require_relative 'excel_functions/more_than_or_equal'
require_relative 'excel_functions/not_equal'

require_relative 'excel_functions/excel_if'
require_relative 'excel_functions/iferror'

# Basic arithmetic
require_relative 'excel_functions/add'
require_relative 'excel_functions/subtract'
require_relative 'excel_functions/multiply'
require_relative 'excel_functions/divide'
require_relative 'excel_functions/power'
require_relative 'excel_functions/negative'

# More advanced arithmetic
require_relative 'excel_functions/abs'
require_relative 'excel_functions/and'
require_relative 'excel_functions/mod'

# Array arithmetic
require_relative 'excel_functions/sum'
require_relative 'excel_functions/average'
require_relative 'excel_functions/max'
require_relative 'excel_functions/min'
require_relative 'excel_functions/subtotal'
require_relative 'excel_functions/count'
require_relative 'excel_functions/counta'
require_relative 'excel_functions/sumif'
require_relative 'excel_functions/sumifs'
require_relative 'excel_functions/sumproduct'

# Lookup functions
require_relative 'excel_functions/index'
require_relative 'excel_functions/excel_match'
require_relative 'excel_functions/vlookup'
require_relative 'excel_functions/choose'
require_relative 'excel_functions/large'

# Financial functions
require_relative 'excel_functions/pmt'

# Geometry functions
require_relative 'excel_functions/cosh'

# String functions
require_relative 'excel_functions/string_join'
require_relative 'excel_functions/find'
require_relative 'excel_functions/left'
require_relative 'excel_functions/right'

# Rounding functions
require_relative 'excel_functions/round'
require_relative 'excel_functions/roundup'
require_relative 'excel_functions/rounddown'
require_relative 'excel_functions/int'

# Other functions


require_relative 'excel_functions/cell'

require_relative 'excel_functions/trim'

require_relative 'excel_functions/mid'

require_relative 'excel_functions/pv'

require_relative 'excel_functions/text'

require_relative 'excel_functions/hlookup'

require_relative 'excel_functions/log'

require_relative 'excel_functions/mmult'

require_relative 'excel_functions/rank'

require_relative 'excel_functions/isnumber'

require_relative 'excel_functions/len'

require_relative 'excel_functions/substitute'
