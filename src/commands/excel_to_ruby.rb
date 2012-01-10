require 'fileutils'
require_relative '../util'
require_relative '../excel'
require_relative '../extract'
require_relative '../rewrite'
require_relative '../simplify'
require_relative '../compile'

class ExcelToRuby
  
  attr_accessor :excel_file, :output_directory, :xml_dir
  
  def go!
    sort_out_output_directories
    unzip_excel
    process_workbook
    extract_worksheets
    Process.waitall
    rewrite_worksheets
    Process.waitall
    simplify_worksheets
    Process.waitall
    compile_worksheets
    Process.waitall
  end
  
  def sort_out_output_directories
    self.excel_file = File.expand_path(excel_file)
    self.output_directory = File.expand_path(output_directory)
    FileUtils.mkdir_p(File.join(output_directory,'intermediate'))
    FileUtils.mkdir_p(File.join(output_directory,'ruby'))
  end
  
  def unzip_excel
    self.xml_dir = File.join(output_directory,'xml')
    puts `unzip -uo '#{excel_file}' -d '#{xml_dir}'`
  end

  def process_workbook    
    extract ExtractSharedStrings, 'sharedStrings.xml', 'shared_strings'
    
    extract ExtractNamedReferences, 'workbook.xml', 'named_references'
    rewrite RewriteFormulaeToAst, 'named_references', 'named_references.ast'
    
    extract ExtractWorksheetNames, 'workbook.xml', 'worksheet_names_without_filenames'
    extract ExtractRelationships, File.join('_rels','workbook.xml.rels'), 'workbook_relationships'
    rewrite RewriteWorksheetNames, 'worksheet_names_without_filenames', 'workbook_relationships', 'worksheet_names'
    extract_dimensions_from_worksheets
  end
  
  # Extracts each worksheets values and formulas
  def extract_worksheets
    worksheets do |name,xml_filename|
      fork do
        $0 = "ruby initial extract #{name}"
        initial_extract_from_worksheet(name,xml_filename)
      end
    end
  end

  # Extracts the dimensions of each worksheet and puts them in a single file  
  def extract_dimensions_from_worksheets    
    dimension_file = output('dimensions')
    worksheets do |name,xml_filename|
      dimension_file.write name
      dimension_file.write "\t"
      extract ExtractWorksheetDimensions, File.open(xml_filename,'r'), dimension_file 
    end
    dimension_file.close
  end
  
  def rewrite_worksheets
    worksheets do |name,xml_filename|
      fork do 
        rewrite_row_and_column_references(name,xml_filename)
        rewrite_shared_formulae(name,xml_filename)
        rewrite_array_formulae(name,xml_filename)
      end
    end
  end
  
  def rewrite_row_and_column_references(name,xml_filename)
    dimensions = input('dimensions')
    %w{simple_formulae.ast shared_formulae.ast array_formulae.ast}.each do |file|
      dimensions.rewind
      i = File.open(File.join(output_directory,'intermediate',name,file),'r')
      o = File.open(File.join(output_directory,'intermediate',name,"#{file}-nocols"),'w')
      RewriteWholeRowColumnReferencesToAreas.rewrite(i,name, dimensions, o)
      close(i,o)
    end
    dimensions.close
  end
  
  def rewrite_shared_formulae(name,xml_filename)
    i = File.open(File.join(output_directory,'intermediate',name,'shared_formulae.ast-nocols'),'r')
    o = File.open(File.join(output_directory,'intermediate',name,"shared_formulae-expanded.ast"),'w')
    RewriteSharedFormulae.rewrite(i,o)
    close(i,o)
  end
  
  def rewrite_array_formulae(name,xml_filename)
    i = File.open(File.join(output_directory,'intermediate',name,'array_formulae.ast-nocols'),'r')
    o = File.open(File.join(output_directory,'intermediate',name,"array_formulae-expanded.ast"),'w')
    RewriteArrayFormulae.rewrite(i,o)
    close(i,o)
  end
  
  # Extracts:
  # Values
  # Formulae (simple, shared and array)
  # Rewrites:
  # the formulae to ast
  def initial_extract_from_worksheet(name,xml_filename)
    worksheet_directory = File.join(output_directory,'intermediate',name)
    FileUtils.mkdir_p(worksheet_directory)
    worksheet_xml = File.open(xml_filename,'r')
    { ExtractValues => 'values', 
      ExtractSimpleFormulae => 'simple_formulae',
      ExtractSharedFormulae => 'shared_formulae',
      ExtractArrayFormulae => 'array_formulae'
    }.each do |_klass,output_filename|
      worksheet_xml.rewind
      extract _klass, worksheet_xml, File.join(name,output_filename)
      if _klass == ExtractValues
        rewrite RewriteValuesToAst, File.join(name,output_filename), File.join(name,"#{output_filename}.ast")
      else
        rewrite RewriteFormulaeToAst, File.join(name,output_filename), File.join(name,"#{output_filename}.ast")
      end  
    end
    worksheet_xml.rewind
    extract ExtractWorksheetTableRelationships, worksheet_xml, File.join(name,'table_rids')
    if File.exists?(File.join(xml_dir,'xl','worksheets','_rels',"#{File.basename(xml_filename)}.rels"))
      extract ExtractRelationships, File.join('worksheets','_rels',"#{File.basename(xml_filename)}.rels"), File.join(name,'relationships')
      rewrite RewriteRelationshipIdToFilename, File.join(name,'table_rids'), File.join(name,'relationships'), File.join(name,'table_filenames')
      tables = output(name,'tables')
      table_extractor = ExtractTable.new(name)
      input(name,'table_filenames').lines.each do |line|
        extract table_extractor, File.join('worksheets',line.strip), tables
      end
    else
      FileUtils.touch File.join(output_directory,'intermediate',name,'relationships')
      FileUtils.touch File.join(output_directory,'intermediate',name,'table_filenames')      
      FileUtils.touch File.join(output_directory,'intermediate',name,'tables')      
    end
    close(worksheet_xml)
  end
  
  def simplify_worksheets
    worksheets do |name,xml_filename|
      fork do 
        simplify_worksheet(name,xml_filename)
      end
    end
  end
  
  def simplify_worksheet(name,xml_filename)
    replace ReplaceSharedStrings, File.join(name,'values.ast'), 'shared_strings', File.join(name,"values_no_shared_strings.ast")
    replace ReplaceNamedReferences, File.join(name,'simple_formulae.ast'), name, 'named_references.ast', File.join(name,"simple_formulae_no_named_references.ast")
    replace ReplaceNamedReferences, File.join(name,'shared_formulae-expanded.ast'),  name, 'named_references.ast', File.join(name,"shared_formulae_no_named_references.ast")
    replace ReplaceNamedReferences, File.join(name,'array_formulae-expanded.ast'), name, 'named_references.ast', File.join(name,"array_formulae_no_named_references.ast")
  end
  
  def compile_worksheets
    worksheets do |name,xml_filename|
      fork do 
        compile_worksheet_code(name,xml_filename)
        compile_worksheet_test(name,xml_filename)
      end
    end    
  end
  
  def compile_worksheet_code(name,xml_filename)
    i = input(name,"simple_formulae_no_named_references.ast")
    o = ruby("#{name.downcase}.rb")
    o.puts "# #{name}"
    o.puts
    o.puts "class #{name.capitalize}"
    CompileToRuby.rewrite(i, o)
    o.puts "end"
    close(i,o)
  end

  def compile_worksheet_test(name,xml_filename)
    i = input(name,"values_no_shared_strings.ast")
    o = ruby("test_#{name.downcase}.rb")
    o.puts "# Test for #{name}"
    o.puts  "require 'test/unit'"
    o.puts  "require_relative '../#{name.capitalize}"
    o.puts
    o.puts "class Test#{name.capitalize} < Test::Unit::TestCase"
    CompileToRubyUnitTest.rewrite(i, o)
    o.puts "end"
    close(i,o)
  end

  
  def worksheets
    IO.readlines(File.join(output_directory,'intermediate','worksheet_names')).each do |line|
      name, filename = *line.split("\t")
      filename = File.expand_path(File.join(xml_dir,'xl',filename.strip))
      yield name, filename
    end
  end
  
  def extract(_klass,xml_name,output_name)
    i = xml_name.is_a?(String) ? xml(xml_name) : xml_name
    o = output_name.is_a?(String) ? output(output_name) : output_name
    _klass.extract(i,o)
    if xml_name.is_a?(String)
      close(i)
    end
    if output_name.is_a?(String)
      close(o)
    end
  end
  
  def rewrite(_klass,*args)
    o = output(args.pop)
    inputs = args.map { |name| input(name) }
    _klass.rewrite(*inputs,o)
    close(*inputs,o)
  end
  
  def replace(_klass,*args)
    o = output(args.pop)
    inputs = args.map { |name| input(name) }
    _klass.replace(*inputs,o)
    close(*inputs,o)
  end
  
  def xml(*args)
    File.open(File.join(xml_dir,'xl',*args),'r')
  end
  
  def input(*args)
    File.open(File.join(output_directory,'intermediate',*args),'r')
  end
  
  def output(*args)
    File.open(File.join(output_directory,'intermediate',*args),'w')
  end
  
  def ruby(*args)
    File.open(File.join(output_directory,'ruby',*args),'w')
  end
  
  def close(*args)
    args.map(&:close)
  end
  
end