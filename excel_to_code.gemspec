Gem::Specification.new do |s|
  s.name = "excel_to_code"
  s.version = '0.1.22'
  s.license = "MIT"
  s.add_dependency('rubypeg')
  s.add_dependency('nokogiri','>= 1.5.0')
  s.add_dependency('rspec','>= 2.7.0')
  s.add_dependency('ffi','>= 1.0.11')
  s.required_ruby_version = ">= 1.9.1"
  s.author = "Thomas Counsell, Green on Black Ltd"
  s.email = "tamc@greenonblack.com"
  s.homepage = "http://github.com/tamc/excel_to_code"
  s.platform = Gem::Platform::RUBY
  s.summary = "Converts .xlxs files into pure ruby 1.9 code or pure C code so that they can be executed without excel"
  s.description = File.read(File.join(File.dirname(__FILE__), 'README.md'))
  s.files = ["LICENSE", "README.md","TODO","{src,bin}/**/*"].map{|p| Dir[p]}.flatten
  s.executables = ["excel_to_c","excel_to_ruby"]
  s.require_path = "src"
  s.has_rdoc = false
end
