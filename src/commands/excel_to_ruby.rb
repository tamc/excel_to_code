require 'fileutils'
require_relative '../extract'
require_relative '../rewrite'

class ExcelToRuby
  
  attr_accessor :excel_file, :output_directory, :xml_dir
  
  def go!
    sort_out_output_directories
    unzip_excel
    process_workbook
  end
  
  def sort_out_output_directories
    self.excel_file = File.expand_path(excel_file)
    self.output_directory = File.expand_path(output_directory)
    FileUtils.mkdir(File.join(output_directory,'intermediate'))
  end
  
  def unzip_excel
    self.xml_dir = File.join(output_directory,'xml')
    puts `unzip -uo #{excel_file} -d #{xml_dir}`
  end

  # Extract the shared strings    
  # Extract the sheet names, initialy with relationship references
  # Extract the workbook relationships  
  def process_workbook    
    extract ExtractSharedStrings, 'sharedStrings.xml', 'shared_strings'
    extract ExtractWorksheetNames, 'workbook.xml', 'worksheet_names_without_filenames'
    extract ExtractRelationships, '_rels/workbook.xml.rels', 'workbook_relationships'
    rewrite RewriteWorksheetNames, 'worksheet_names_without_filenames', 'workbook_relationships', 'worksheet_names'
  end
  
  def extract(_klass,xml_name,output_name)
    i = xml(xml_name)
    o = output(output_name)
    _klass.extract(i,o)
    close(i,o)
  end
  
  def rewrite(_klass,*args)
    o = output(args.pop)
    inputs = args.map { |name| input(name) }
    _klass.rewrite(*inputs,o)
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
  
  def close(*args)
    args.map(&:close)
  end
  
end