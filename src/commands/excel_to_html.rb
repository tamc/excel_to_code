# coding: utf-8
require 'fileutils'
require 'logger'
require_relative '../excel_to_code'

# Used to throw normally fatal errors
class ExcelToCodeException < Exception; end
class VersionedFileNotFoundException < Exception; end
class XMLFileNotFoundException < Exception; end

class ExcelToHTML
  
  # Required attribute. The source excel file. This must be .xlsx not .xls
  attr_accessor :excel_file
  
  # Optional attribute. The output directory.
  #  If not specified, will be '#{excel_file_name}/c'
  attr_accessor :output_directory
  
  # Optional attribute. The excel file will be translated to xml and stored here.
  # If not specified, will be '#{excel_file_name}/html'
  attr_accessor :xml_directory

  # Optional attribute. The intermediate workings will be stored here.
  # If not specified, will be '#{excel_file_name}/intermediate'
  attr_accessor :intermediate_directory
  
  # Optional attribute. Boolean.
  #   * true - the intermediate files are not written to disk (requires a lot of memory)
  #   * false - the intermediate files are written to disk (default, easier to debug)
  attr_accessor :run_in_memory
  
  # This is the log file, if set it needs to respond to the same methods as the standard logger library
  attr_accessor :log

  def set_defaults
    raise ExcelToCodeException.new("No excel file has been specified") unless excel_file
    
    self.output_directory ||= File.join(File.dirname(excel_file),File.basename(excel_file,".*"),"html")
    self.xml_directory ||= File.join(File.dirname(excel_file),File.basename(excel_file,".*"),'xml')
    self.intermediate_directory ||= File.join(File.dirname(excel_file),File.basename(excel_file,".*"),'intermediate')
    
    # Make sure the relevant directories exist
    self.excel_file = File.expand_path(excel_file)
    self.output_directory = File.expand_path(output_directory)
    
    # Set up our log file
    self.log ||= Logger.new(STDOUT)
  end
  
  def go!
    # This sorts out the settings
    set_defaults
    
    # These turn the excel into a more accesible format
    sort_out_output_directories
    unzip_excel
    
    # These get all the information out of the excel and put
    # into a series of plain text files
    extract_data_from_workbook
    extract_data_from_worksheets

    # Combines the different formulae types
    rewrite_worksheets

    # At the moment, this just replaces shared strings
    simplify_worksheets

    # This puts the results into a spreadsheet
    write_out_excel_as_html
    
    log.info "The generated html is available in #{File.join(output_directory)}"
  end
  
  # Creates any directories that are needed
  def sort_out_output_directories    
    FileUtils.mkdir_p(output_directory)
    FileUtils.mkdir_p(xml_directory)
    FileUtils.mkdir_p(intermediate_directory) unless run_in_memory
  end
  
  # FIXME: Replace these with pure ruby versions?
  def unzip_excel
    log.info `rm -fr '#{xml_directory}'` # Force delete
    log.info `unzip '#{excel_file}' -d '#{xml_directory}'` # If don't force delete, make sure that force the zip to overwrite old files 
  end
  
  # The excel workbook.xml and allied relationship files knows about
  # shared strings, named references and the actual human readable
  # names of each of the worksheets. 
  #
  # In this method we also loop through each of the individual 
  # worksheet files to work out their dimensions
  def extract_data_from_workbook
    extract_shared_strings
    extract_named_references
    extract_worksheet_names
    extract_dimensions_from_worksheets
  end
  
  # Excel keeps a central file of strings that appear in worksheet cells
  def extract_shared_strings
    extract ExtractSharedStrings, 'sharedStrings.xml', 'Shared strings'
  end
  
  # Excel keeps a central list of named references. This includes those
  # that are local to a specific worksheet.
  def extract_named_references
    extract ExtractNamedReferences, 'workbook.xml', 'Named references'
    apply_rewrite RewriteFormulaeToAst, 'Named references'
    replace ReplaceRangesWithArrayLiterals, 'Named references', 'Named references'
  end

  # Excel keeps a list of worksheet names. To get the mapping between
  # human and computer name  correct we have to look in the workbook 
  # relationships files. We also need to mangle the name into something
  # that will work ok as a filesystem or program name
  def extract_worksheet_names
    extract ExtractWorksheetNames, 'workbook.xml', 'Worksheet names'
    extract ExtractRelationships, File.join('_rels','workbook.xml.rels'), 'Workbook relationships'
    rewrite RewriteWorksheetNames, 'Worksheet names', 'Workbook relationships', 'Worksheet names'
  end

  # We want a central list of the maximum extent of each worksheet
  # so that we can convert column (e.g., C:F) and row (e.g., 13:18)
  # references into equivalent area references (e.g., C1:F30)
  def extract_dimensions_from_worksheets 
    log.info "Starting to extract dimensions from worksheets"  
    dimension_file = intermediate('Worksheet dimensions')
    extractor = ExtractWorksheetDimensions.new
    worksheets do |name, xml_filename|
      log.info "Extracting dimensions for #{name}"
      dimension_file.write name
      dimension_file.write "\t"
      
      extractor.extract(xml(xml_filename), dimension_file)
      close(xml_filename)
    end
    close(dimension_file)
  end
  
  # For each worksheet, this makes four passes through the xml
  # 1. Extract the values of each cell
  # 2. Extract all the cells which are simple formulae
  # 3. Extract all the cells which use shared formulae
  # 4. Extract all the cells which are part of array formulae
  def extract_data_from_worksheets
    worksheets do |name, xml_filename|
      
      extract ExtractValues, xml_filename, [name, 'Values']
      apply_rewrite RewriteValuesToAst, [name, 'Values']
      
      extract ExtractSimpleFormulae, xml_filename, [name, 'Formulae (simple)']
      apply_rewrite RewriteFormulaeToAst, [name, 'Formulae (simple)']

      extract ExtractSharedFormulae, xml_filename, [name, 'Formulae (shared)']
      apply_rewrite RewriteFormulaeToAst, [name, 'Formulae (shared)']

      extract ExtractSharedFormulaeTargets, xml_filename, [name, 'Formulae (shared targets)']

      extract ExtractArrayFormulae, xml_filename, [name, 'Formulae (array)']
      apply_rewrite RewriteFormulaeToAst, [name, 'Formulae (array)']
      
    end
  end
  
  def rewrite_worksheets
    worksheets do |name,xml_filename|
      log.info "Rewriting worksheet #{name}"
      rewrite_shared_formulae(name,xml_filename)
      rewrite_array_formulae(name,xml_filename)
      combine_formulae_files(name,xml_filename)
    end
  end
  
  def rewrite_shared_formulae(name,xml_filename)
    rewrite RewriteSharedFormulae, [name, 'Formulae (shared)'], [name, 'Formulae (shared targets)'], [name, 'Formulae (shared)']
  end
  
  def rewrite_array_formulae(name,xml_filename)
    r = ReplaceNamedReferences.new
    r.sheet_name = name
    replace r, [name, 'Formulae (array)'], 'Named references', [name, 'Formulae (array)']

    r = ReplaceTableReferences.new
    r.sheet_name = name    
    replace r,                              [name, 'Formulae (array)'], "Workbook tables", [name, 'Formulae (array)']
    replace SimplifyArithmetic,             [name, 'Formulae (array)'], [name, 'Formulae (array)']
    replace ReplaceRangesWithArrayLiterals, [name, 'Formulae (array)'], [name, 'Formulae (array)']
    apply_rewrite RewriteArrayFormulaeToArrays,   [name, 'Formulae (array)']
    apply_rewrite RewriteArrayFormulae,           [name, 'Formulae (array)']
  end
  
  def combine_formulae_files(name,xml_filename)
    combiner = RewriteMergeFormulaeAndValues.new
    rewrite combiner, [name, 'Values'], [name, 'Formulae (shared)'], [name, 'Formulae (array)'], [name, 'Formulae (simple)'], [name, 'Formulae']
  end
  
  def simplify_worksheets
    worksheets do |name,xml_filename|
      replace ReplaceSharedStrings, [name, 'Values'], 'Shared strings', File.join(name, 'Values')
    end
  end
      
  # Returns a hash of named references, and the ast of their links
  # where the named reference is global the key will be a string of
  # its name and case sensitive.
  # where the named reference is coped to a worksheet, the key will be
  # a two element array. The first element will be the sheet name. The
  # second will be the name. 
  def named_references
    return @named_references if @named_references
    @named_references = {}
    i = input('Named references')
    i.lines.each do |line|
      sheet, name, ref = *line.split("\t")
      key = sheet.size != 0 ? [sheet, name] : name
      @named_references[key] = eval(ref)
    end
    close(i)
    @named_references
  end

  # Write out the html
  def write_out_excel_as_html  
    #
    # Create an index file that just redirects to the first sheet
    main_sheet_name = worksheets.first.first
    File.open(File.join(output_directory,'index.html'),'w') do |f|
      f.puts <<-END
        <html>
        <meta http-equiv="refresh" content="0; url=#{main_sheet_name}.html" />
        <link rel="canonical" href="#{main_sheet_name}.html"/>
        </html>
      END
    end

    # Copy across the stylesheets and the javascripts
    %w{ jquery.min.js application.css application.js }.each do |file|
      `cp #{File.join(File.dirname(__FILE__),'..','compile','html',file)} #{File.join(output_directory,file)}`
    end

    # Now go through and create a web page for each worksheet
    c = CompileToHTML.new
    c.worksheet_dimensions = input('Worksheet dimensions')
    c.formulae = all_formulae
    c.values = all_values

    worksheets do |name, sheet|
      log.info "Writing out the html for #{name}"
      o = output("#{name}.html")
      c.rewrite(name, o)
      close(o)
    end

  end

  # UTILITY FUNCTIONS
    
  def all_formulae
    references = {}
    worksheets do |name,xml_filename|
      r = references[name] = {}
      i = input([name,'Formulae'])
      i.lines do |line|
        line =~ /^(.*?)\t(.*)$/
        ref, ast = $1, $2
        r[ref] = eval(ast)
      end
    end 
    references
  end

  def all_values
    references = {}
    worksheets do |name,xml_filename|
      r = references[name] = {}
      i = input([name,'Values'])
      i.lines do |line|
        line =~ /^(.*?)\t(.*)$/
        ref, ast = $1, $2
        r[ref] = eval(ast)
      end
    end 
    references
  end
    
  def worksheets(&block)
    unless @worksheet_filenames
      worksheet_names = input('Worksheet names')
      @worksheet_filenames = worksheet_names.lines.map do |line|
        name, filename = *line.split("\t")
        [name, filename.strip]
      end
      close(worksheet_names)
    end
   
   if block 
      @worksheet_filenames.each do |name, filename|
        block.call(name, filename)
      end
   end
   @worksheet_filenames
  end
    
  def extract(klass,xml_name,output_name)
    log.debug "Started using #{klass} to extract xml: #{xml_name} to #{output_name}"
    
    i = xml(xml_name)
    o = intermediate(output_name)
    klass.extract(i,o)
    close(i,o)

    log.info "Finished using #{klass} to extract xml: #{xml_name} to #{output_name}"
  end
  
  def apply_rewrite(klass,filename)
    rewrite klass, filename, filename
  end
  
  def rewrite(klass, *args)
    execute klass, :rewrite, *args
  end
  
  def replace(klass, *args)
    execute klass, :replace, *args
  end
  
  def execute(klass, method, *args)
    log.debug "Started executing #{klass}.#{method} with #{args.inspect}"
    inputs = args[0..-2].map { |name| input(name) }
    output = intermediate(args.last)
    klass.send(method,*inputs,output)
    close(*inputs,output)
    log.info "Finished executing #{klass}.#{method} with #{args.inspect}"
  end
  
  def xml(*args)
    args.flatten!
    filename = File.join(xml_directory,'xl',*args)
    if File.exists?(filename)
      File.open(filename,'r')
    else
      log.warn("#{filename} does not exist in xml(#{args.inspect}), using blank instead")
      StringIO.new
    end
  end
  
  def input(*args)
    args.flatten!
    filename = versioned_filename_read(intermediate_directory,*args)
    if run_in_memory
      existing_file = @files[filename]
      if existing_file
        StringIO.new(existing_file.string,'r')
      else
        log.warn("#{filename} does not exist in input(#{args.inspect}), using blank instead")
        StringIO.new
      end
    else
      if File.exists?(filename)
        File.open(filename,'r')
      else
        log.warn("#{filename} does not exist in input(#{args.inspect}), using blank instead")
        StringIO.new
      end
    end
  end
  
  def intermediate(*args)
    args.flatten!
    filename = versioned_filename_write(intermediate_directory,*args)
    if run_in_memory
      @files ||= {}
      @files[filename] = StringIO.new("",'w')
    else
      FileUtils.mkdir_p(File.dirname(filename))
      File.open(filename,'w')
    end
  end
  
  def output(*args)
    args.flatten!
    File.open(File.join(output_directory,*args),'w')
  end
  
  def close(*args)
    args.map do |f|
      next if f.is_a?(StringIO)
      next if f.is_a?(String)
      f.close
    end
  end
  
  def versioned_filename_read(*args)
    @versioned_filenames ||= {}
    standardised_name = standardise_name(args)
    counter = @versioned_filenames[standardised_name]
    filename_with_counter counter, args
  end
  
  def versioned_filename_write(*args)
    @versioned_filenames ||= {}
    standardised_name = standardise_name(args)
    if @versioned_filenames.has_key?(standardised_name)
      counter =  @versioned_filenames[standardised_name] + 1
    else
      counter = 0
    end
    @versioned_filenames[standardised_name] = counter
    filename_with_counter(counter, args)
  end
  
  def filename_with_counter(counter, args)
    counter ||= 0
    last_name = args.last
    last_name = last_name + sprintf(" %03d", counter)
    File.join(*args[0..-2], last_name)    
  end  
  
  def standardise_name(*args)
    File.expand_path(File.join(args))
  end
  
end
