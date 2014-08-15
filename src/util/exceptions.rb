class ExcelToCodeException < Exception; end

# Thrown when some valid bit of Excel has not been implemented in this converter
class NotSupportedException < ExcelToCodeException; end
