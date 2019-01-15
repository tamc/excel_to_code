require_relative 'src/version'

Gem::Specification.new do |s|
  s.name = "excel_to_code"
  s.version = ExcelToCode.version
  s.license = "MIT"
  s.add_runtime_dependency 'rubypeg', '~> 0', '>= 0.0.4'
  s.add_runtime_dependency 'rspec', '~> 3.7'
  s.add_runtime_dependency 'ffi', '~> 1.9', '>= 1.9.18'
  s.add_runtime_dependency 'ox', '~> 2.8', '>= 2.8.2'
  s.add_runtime_dependency 'minitest', '~> 5.11', '>= 5.11.1'
  s.add_development_dependency 'rake', '~> 12'
  s.required_ruby_version = ">= 2.3.0"
  s.author = "Thomas Counsell, Green on Black Ltd"
  s.email = "tamc@greenonblack.com"
  s.homepage = "http://github.com/tamc/excel_to_code"
  s.platform = Gem::Platform::RUBY
  s.summary = "Convert many .xlsx and .xlsm files into equivalent Ruby or C code that can be executed without Excel"
  s.description = File.read(File.join(File.dirname(__FILE__), 'README.md'))
  s.files = ["LICENSE", "README.md","TODO","{src,bin}/**/*"].map{|p| Dir[p]}.flatten
  s.executables = ["excel_to_c","excel_to_ruby"]
  s.require_path = "src"
end
