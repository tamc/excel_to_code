# coding: utf-8
# Used to throw normally fatal errors
class ExcelToCodeException < Exception; end
class XMLFileNotFoundException < Exception; end

require 'fileutils'
require 'logger'
require 'tmpdir'
require_relative '../excel_to_code'

# FIXME: Correct case for all worksheet references
# FIXME: Correct case and $ stripping from all cell references
# FIXME: Replacing with c compatible names everywhere

class ExcelToX
  
  # Required attribute. The source excel file. This must be .xlsx not .xls
  attr_accessor :excel_file
  
  # Optional attribute. The output directory.
  #  If not specified, will be '#{excel_file_name}/c'
  attr_accessor :output_directory
  
  # Optional attribute. The name of the resulting ruby or c file and ruby or ruby ffi module name. Defaults to excelspreadsheet
  attr_accessor :output_name

  # Optional attribute. The excel file will be translated to xml and stored here.
  # If not specified, will be '#{excel_file_name}/xml'
  attr_accessor :xml_directory
  
  # Optional attribute. Specifies which cells have setters created in the c code so their values can be altered at runtime.
  # It is a hash. The keys are the sheet names. The values are either the symbol :all to specify that all cells on that sheet 
  # should be setable, or an array of cell names on that sheet that should be settable (e.g., A1)
  attr_accessor :cells_that_can_be_set_at_runtime

  # Optional attribute. Specifies which named references to be turned into setters. 
  #
  # NB: Named references are assumed to include table names.
  #
  # Should be an array of strings. Each string is a named reference. Case sensitive.
  # To specify a named reference scoped to a worksheet, use ['worksheet', 'named reference'] instead
  # of a string.
  #
  # Alternatively, can se to :where_possible to create setters for named references that point to setable cells
  #
  # Alternatively, can specify a block, in which case each named reference in the workbook will be yielded to it
  # and if the bock returns true, it will be made settable
  #
  # Each named reference then has a function in the resulting C code of the form
  # void set_named_reference_mangled_into_a_c_function(ExcelValue newValue)
  #
  # By default no named references are output
  attr_accessor :named_references_that_can_be_set_at_runtime
  
  # Optional attribute. Specifies which cells must appear in the final generated code.
  # The default is that all cells in the original spreadsheet appear in the final code.
  #
  # If specified, then any cells that are not:
  #    * specified
  #    * required to calculate one of the cells that is specified
  #    * specified as a cell that can be set at runtime
  # may be excluded from the final generated code.
  #
  # It is a hash. The keys are the sheet names. The values are either the symbol :all to specify that all cells on that sheet 
  # should be lept, or an array of cell names on that sheet that should be kept (e.g., A1)
  attr_accessor :cells_to_keep
 
  # Optional attribute. Specifies which named references should be included in the output
  #
  # NB: Named references are assumed to include table names.
  #
  # Should be an array of strings. Each string is a named reference. Case sensitive.
  #
  # To specify a named reference scoped to a worksheet, use ['worksheet', 'named reference'] instead
  # of a string.
  #
  # Alternatively, can specify :all to keep all named references
  #
  # Alternatively, can specify a block, in which case each named reference in the workbook will be yielded to it
  # and if the bock returns true, it will be kept in the output and if false it may be optimised out of the output
  #
  # Each named reference then has a function in the resulting C code of the form
  # ExcelValue named_reference_mangled_into_a_c_function()
  #
  # By default, no named references are output
  attr_accessor :named_references_to_keep
  
  # Optional attribute. Boolean. Not relevant to all types of code output
  #   * true - the generated c code is compiled
  #   * false - the generated c code is not compiled (default, unless actuall_run_tests is specified as true)
  attr_accessor :actually_compile_code

  # Optional attribute. Boolean. 
  #   * true - the generated tests are run
  #   * false (default) - the generated tests are not run
  attr_accessor :actually_run_tests
  
  # This is the log file, if set it needs to respond to the same methods as the standard logger library
  attr_accessor :log

  # Optional attribute. Boolean.
  #   * true (default) - empty cells and zeros are treated as being equivalent in tests. Numbers greater then 1 are only expected to match with assert_in_epsilon, numbers less than 1 are only expected to match with assert_in_delta
  #   * false - empty cells and zeros are treated as being different in tests. Numbers must match to full accuracy.
  attr_accessor :sloppy_tests

  # Optional attribute, Boolean.
  #   * true (default) - the compiler attempts to inline any calculation that is done in another cell, but only referred to by this cell. This should increase performance
  #   * false - the compiler leaves calculations in their original cells expanded. This may make debugging easier
  attr_accessor :should_inline_formulae_that_are_only_used_once
  
  # Optional attribute, Boolean.
  #   * true (default) - the compiler attempts to extract bits of calculation that appear in more than one formula into separate methods. This should increase performance
  #   * false - the compiler leaves calculations fully expanded. This may make debugging easier
  attr_accessor :extract_repeated_parts_of_formulae


  # Optional attribute, Array. Default nil
  # This is used to help debug large spreadsheets that aren't working correctly.
  # If set to the name of a worksheet then ONLY that worksheet will be run through the 
  # optimisation and simplification code. Will also override cells_to_keep to keep all
  # cells on tha sheet and nothing else.
  attr_accessor :isolate

  # This is the main method. Once all the above attributes have been set, it should be called to actually do the work.
  def go!
    # This sorts out the settings
    set_defaults

    log.info "Excel to Code version #{ExcelToCode.version}\n\n"
    
    # These turn the excel into xml on disk
    sort_out_output_directories
    unzip_excel
    
    # These gets the named references, worksheet names and shared strings out of the excel
    extract_data_from_workbook
    
    # This gets all the formulae, values and tables out of the worksheets
    extract_data_from_worksheets
    
    # This checks that the user inputs of which cells to keep are in the right
    # format and refer to sheets and references that actually exist
    clean_cells_that_can_be_set_at_runtime
    clean_cells_to_keep
    convert_named_references_into_simple_form
    clean_named_references_to_keep
    clean_named_references_that_can_be_set_at_runtime

    # This is an early check that the functions in the extracted data have 
    # all got an implementation in, at least, the ruby code
    check_all_functions_implemented

    # This turns named references that are specified as getters and setters
    # into a series of required cell references
    transfer_named_references_to_keep_into_cells_to_keep
    transfer_named_references_that_can_be_set_at_runtime_into_cells_that_can_be_set_at_runtime
    
    # These perform some translations to tsimplify the excel
    # Including:
    # * Turning row and column references (e.g., A:A) to areas, based on the size of the worksheet
    # * Turning range references (e.g., A1:B2) into array litterals (e.g., {A1,B1;A2,B2})
    # * Turning shared formulae into a series of conventional formulae
    # * Turning array formulae into a series of conventional formulae
    # * Mergining all the different types of formulae and values into a single hash
    rewrite_values_to_remove_shared_strings
    rewrite_row_and_column_references
    rewrite_shared_formulae_into_normal_formulae
    rewrite_array_formulae
    combine_formulae_types
    
    # These perform a series of transformations to the information
    # with the intent of removing any redundant calculations that are in the excel.
    # Replacing shared strings and named references with their actual values, tidying arithmetic
    simplify_arithmetic
    simplify

    # If nothing has been specified in named_references_that_can_be_set_at_runtime 
    # or in cells_that_can_be_set_at_runtime, then we assume that
    # all value cells should be settable if they are referenced by
    # any other forumla.
    ensure_there_is_a_good_set_of_cells_that_can_be_set_at_runtime

    # If named_reference_that_have_been_set_at_runtime is given the :where_possible switch
    # then will take a look at which named_references only refer to cells that have been
    # specifed or judged as settable
    work_out_which_named_references_can_be_set_at_runtime

    # Slims down the named references we keep track of to just the ones that should
    # appear in the generated code: basically those that are specifed as being gettable
    # or specified or judged to be settable.
    filter_named_references

    replace_formulae_with_their_results
    inline_formulae_that_are_only_used_once if should_inline_formulae_that_are_only_used_once
    remove_any_cells_not_needed_for_outputs
    separate_formulae_elements if extract_repeated_parts_of_formulae
    replace_values_with_constants
    create_sorted_references_to_test

    # This actually creates the code (implemented in subclasses)
    write_code
    
    # These compile and run the code version of the excel (implemented in subclasses)
    compile_code
    run_tests

    cleanup

    log.info "The generated code is available in #{File.join(output_directory)}"
  end
  
  # If an attribute hasn't been specified, specifies a good default value here.
  def set_defaults
    raise ExcelToCodeException.new("No excel file has been specified") unless excel_file
    
    self.output_directory ||= Dir.pwd
    unless self.xml_directory
      self.xml_directory ||= Dir.mktmpdir
      @delete_xml_directory_at_end = true
    end
    
    self.output_name ||= "Excelspreadsheet"
    
    self.cells_that_can_be_set_at_runtime ||= {}
    
    # Make sure the relevant directories exist
    self.excel_file = File.expand_path(excel_file)
    self.output_directory = File.expand_path(output_directory)
    
    # Set up our log file
    unless self.log
      self.log = Logger.new(STDOUT)
      log.formatter = proc do |severity, datetime, progname, msg|
        case severity
        when "FATAL"; "\033[41m#{datetime.strftime("%H:%M")}\t#{msg}\033[0m\n"
        when "ERROR"; "\033[41m#{datetime.strftime("%H:%M")}\t#{msg}\033[0m\n"
        when "WARN"; "\033[31m#{datetime.strftime("%H:%M")}\t#{msg}\033[0m\n"
        when "INFO"; "\033[34m#{datetime.strftime("%H:%M")}\t#{msg}\033[0m\n"
        else; "#{datetime.strftime("%H:%M")}\t#{msg}\n"
        end
      end
    end

    # By default, tests allow empty cells and zeros to be treated as equivalent, and numbers only have to match to a 0.001 epsilon (if expected>1) or 0.001 delta (if expected<1)
    self.sloppy_tests ||= true

    # Setting this to false may make it easier to figure out errors
    self.extract_repeated_parts_of_formulae = true if @extract_repeated_parts_of_formulae == nil
    self.should_inline_formulae_that_are_only_used_once = true if @should_inline_formulae_that_are_only_used_once == nil

    # This setting is used for debugging, and makes the system only do the conversion on a subset of the worksheets
    if self.isolate
      self.isolate = [self.isolate] unless self.isolate.is_a?(Array)
      self.cells_to_keep ||= {}
      self.isolate.each do |sheet|
        self.cells_to_keep[sheet.to_s] = :all
      end
      self.isolate = self.isolate.map { |s| s.to_sym }
      log.warn "Isolating #{@isolate} worksheet(s). No other sheets will be converted"
    end
  end

  # Creates any directories that are needed
  def sort_out_output_directories    
    FileUtils.mkdir_p(output_directory)
    FileUtils.mkdir_p(xml_directory)
  end
  
  # FIXME: Replace these with pure ruby versions?
  def unzip_excel
    log.info "Removing any old xml #{`rm -fr '#{xml_directory}'`}" # Force delete
    log.info "Unziping excel into xml #{`unzip -q '#{excel_file}' -d '#{xml_directory}'`}" # If don't force delete, make sure that force the zip to overwrite old files 
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
  end

  # @shared_strings is an array of strings
  def extract_shared_strings
    log.info "Extracting shared strings"
    # Excel keeps a central file of strings that appear in worksheet cells
    xml('sharedStrings.xml') do |i|
      @shared_strings = ExtractSharedStrings.extract(i)
    end
  end
  
  # Excel keeps a central list of named references. This includes those
  # that are local to a specific worksheet.
  # They are put in a @named_references hash
  # The hash value is the ast for the reference
  # The hash key is either [sheet, name] or name
  # Note that the sheet and the name are always stored lowercase
  def extract_named_references
    log.info "Extracting named references"
    # First we get the references in raw form
    xml('workbook.xml') do |i|
      @named_references = ExtractNamedReferences.extract(i)
    end
    # Then we parse them
    @named_references.each do |name, reference|
      begin
        parsed = CachingFormulaParser.parse(reference)
        if parsed
          @named_references[name] = parsed
        else
          $stderr.puts "Named reference #{name} #{reference} not parsed"
          exit
        end
      rescue Exception
        $stderr.puts "Named reference #{name} #{reference} not parsed"
        raise
      end
    end

  end

  # Named references can be simple cell references, 
  # or they can be ranges, or errors, or table references
  # this function converts all the different types into
  # arrays of cell references
  def convert_named_references_into_simple_form
    # Replace A$1:B2 with [A1, A2, B1, B2]
    @replace_ranges_with_array_literals_replacer ||= ReplaceRangesWithArrayLiteralsAst.new
    table_reference_replacer = ReplaceTableReferenceAst.new(@tables)

    @named_references.each do |name, reference|
      reference = table_reference_replacer.map(reference)
      reference = @replace_ranges_with_array_literals_replacer.map(reference)
      @named_references[name] = reference
    end

  end

  # Excel keeps a list of worksheet names. To get the mapping between
  # human and computer name  correct we have to look in the workbook 
  # relationships files. We also need to mangle the name into something
  # that will work ok as a filesystem or program name
  def extract_worksheet_names
    log.info "Extracting worksheet names"
    
    worksheet_rids = {}

    xml('workbook.xml') do |i|
      worksheet_rids = ExtractWorksheetNames.extract(i) # {'worksheet_name' => 'rId3' ...}
    end
    
    xml_for_rids = {}
    xml('_rels','workbook.xml.rels') do |i|
      xml_for_rids = ExtractRelationships.extract(i) #{ 'rId3' => "worlsheets/sheet1.xml" }
    end

    @worksheet_xmls = {}
    worksheet_rids.each do |name, rid|
      worksheet_xml = xml_for_rids[rid]
      if worksheet_xml =~ /^worksheets/i # This gets rid of things that look like worksheets but aren't (e.g., chart sheets)
        @worksheet_xmls[name.to_sym] = worksheet_xml
      end
    end
    # FIXME: Extract this and put it at the end ?
    @worksheet_c_names = {}
    worksheet_rids.keys.each do |excel_worksheet_name|
      @worksheet_c_names[excel_worksheet_name] = @worksheet_c_names[excel_worksheet_name.to_sym] = c_name_for(excel_worksheet_name)
    end
  end

  def c_name_for(name)
    name = name.to_s
    @c_names_assigned ||= {}
    return @c_names_assigned.invert.fetch(name) if @c_names_assigned.has_value?(name)
    c_name = name.downcase.gsub(/[^a-z0-9]+/,'_') # Make it lowercase, replace anything that isn't a-z or 0-9 with underscores
    c_name = "s"+c_name if c_name[0] !~ /[a-z]/ # Can't start with a number. If it does, but an 's' in front (so 2010 -> s2010)
    c_name = c_name + "2" if @c_names_assigned.has_key?(c_name) # Add a number at the end if the c_name has already been used
    c_name.succ! while @c_names_assigned.has_key?(c_name)
    @c_names_assigned[c_name] = name
    c_name
  end

  # Make sure that sheet names are symbols FIXME: Case ?
  # Make sure that all the cell names are upcase symbols and don't have any $ in them
  def clean_cells_that_can_be_set_at_runtime
    return unless cells_that_can_be_set_at_runtime.is_a?(Hash)

    # Make sure sheet names are symbols
    cells_that_can_be_set_at_runtime.keys.each do |sheet|
      next if sheet.is_a?(Symbol)
      cells_that_can_be_set_at_runtime[sheet.to_sym] = cells_that_can_be_set_at_runtime.delete(sheet)
    end

    # Make sure the sheets actually exist
    cells_that_can_be_set_at_runtime.keys.each do |sheet|
      next if @worksheet_xmls.has_key?(sheet)
      log.error "Cells that can be set at runtime includes #{sheet.inspect} but could not be found in workbook: #{@worksheet_xmls.keys.inspect}"
      exit
    end

    # Make sure references are of the form A1, not a1 or A$1
    cells_that_can_be_set_at_runtime.keys.each do |sheet|
      next unless cells_that_can_be_set_at_runtime[sheet].is_a?(Array)
      cells_that_can_be_set_at_runtime[sheet] = cells_that_can_be_set_at_runtime[sheet].map do |reference| 
        reference.gsub('$','').upcase.to_sym
      end
    end
  end

  # Make sure that sheet names are symbols FIXME: Case ?
  # Make sure that all the cell names are upcase symbols and don't have any $ in them
  def clean_cells_to_keep
    return unless cells_to_keep
    
    # Make sure sheet names are symbols
    cells_to_keep.keys.each do |sheet|
      next if sheet.is_a?(Symbol)
      cells_to_keep[sheet.to_sym] = cells_to_keep.delete(sheet)
    end
    
    # Make sure the sheets actually exist
    cells_to_keep.keys.each do |sheet|
      next if @worksheet_xmls.has_key?(sheet)
      log.error "Cells to keep includes #{sheet.inspect} but could not be found in workbook: #{@worksheet_xmls.keys.inspect}"
      exit
    end

    # Make sure references are of the form A1, not a1 or A$1
    cells_to_keep.keys.each do |sheet|
      next unless cells_to_keep[sheet].is_a?(Array)
      cells_to_keep[sheet] = cells_to_keep[sheet].map { |reference| reference.gsub('$','').upcase.to_sym }
    end
  end  

  # Make sure named_references_to_keep are lowercase symbols
  def clean_named_references_to_keep
    # Named references_to_keep can be passed a block, in which case this loops
    # through offering up the named references. If the block returns true then
    # the named reference is kept
    if named_references_to_keep == :all
      @named_references_to_keep = @named_references.keys.concat(@table_areas.keys)
    end

    if named_references_to_keep.is_a?(Proc)
      new_named_references_to_keep = @named_references.keys.select do |named_reference|
        named_references_to_keep.call(named_reference)
      end
      table_references_to_keep = @table_areas.keys.select do |table_name|
        named_references_to_keep.call(table_name)
      end

      @named_references_to_keep = new_named_references_to_keep.concat(table_references_to_keep)
    end

    return unless named_references_to_keep.is_a?(Array)
    named_references_to_keep.map! { |named_reference| named_reference.downcase.to_sym }

    # Now we need to check the user specified named references actually exist
    named_references_to_keep.each.with_index do |named_reference, i|
      next if @named_references.has_key?(named_reference) || @table_areas.has_key?(named_reference)
      $stderr.puts "Named reference #{named_reference.inspect} in named_references_to_keep has not been found in the spreadsheet: #{@named_references.keys.inspect}"
      exit
    end
  end

  # Make sure named_references_that_can_be_set_at_runtime are lowercase symbols
  def clean_named_references_that_can_be_set_at_runtime
    # amed_references_that_can_be_set_at_runtime can be passed a block, in which case this loops
    # through offering up the named references. If the block returns true then
    # the named reference is made settable
    if named_references_that_can_be_set_at_runtime.is_a?(Proc)
      new_named_references_that_can_be_set_at_runtime = @named_references.keys.select do |named_reference|
        named_references_that_can_be_set_at_runtime.call(named_reference)
      end
      table_references_that_can_be_set_at_runtime = @table_areas.keys.select do |table_name|
        named_references_that_can_be_set_at_runtime.call(table_name)
      end
      @named_references_that_can_be_set_at_runtime = new_named_references_that_can_be_set_at_runtime.concat(table_references_that_can_be_set_at_runtime)
    end

    return unless named_references_that_can_be_set_at_runtime.is_a?(Array)
    named_references_that_can_be_set_at_runtime.map! { |named_reference| named_reference.downcase.to_sym }

    # Now we need to check the user specified named references actually exist
    named_references_that_can_be_set_at_runtime.each.with_index do |named_reference, i|
      next if @named_references.has_key?(named_reference) || @table_areas.has_key?(named_reference)
      $stderr.puts "Named reference #{named_reference.inspect} in named_references_that_can_be_set_at_runtime has not been found in the spreadsheet: #{@named_references.keys.inspect}"
      exit
    end
  end

  
  # For each worksheet, extract the useful bits from the excel xml
  def extract_data_from_worksheets
    # All are hashes of the format ["SheetName", "A1"] => [:number, "1"]
    # This one has a series of table references
    extractor = ExtractDataFromWorksheet.new
    
    # Loop through the worksheets
    # FIXME: make xml_filename be the IO object?
    worksheets do |name, xml_filename|

      # This is used in debugging large worksheets to limit 
      # the optimisation to a particular worksheet
      if isolate
        log.info "Only extracting values from #{name}: #{!isolate.include?(name)}"
        extractor.only_extract_values = !isolate.include?(name)
      end

      log.info "Extracting data from #{name}"
      xml(xml_filename) do |input|
        extractor.extract(name, input)
      end
    end
    @values = extractor.values
    @formulae_simple = extractor.formulae_simple
    @formulae_shared = extractor.formulae_shared
    @formulae_shared_targets = extractor.formulae_shared_targets 
    @formulae_array = extractor.formulae_array
    @worksheets_dimensions = extractor.worksheets_dimensions
    @table_rids = extractor.table_rids
    @tables = {}
    @table_areas = {}
    @table_data = {}
    extract_tables
  end
  
  # To extract a table we need to look in the worksheet for table references
  # then we look in the relationships file for the filename that matches that
  # reference and contains the table data. Then we consolidate all the data
  # from individual table files into a single table file for the worksheet.
  def extract_tables
    log.info "Extracting Tables"
    @table_rids.each do |worksheet_name, array_of_table_rids|
      xml_filename = @worksheet_xmls[worksheet_name]
      xml_for_rids = {}

      # Load the relationship file
      xml(File.join('worksheets','_rels',"#{File.basename(xml_filename)}.rels")) do |i|
        xml_for_rids = ExtractRelationships.extract(i)
      end
      
      # Then extract the individual tables
      array_of_table_rids.each do |rid| 
        xml(File.join('worksheets', xml_for_rids[rid])) do |i|
          ExtractTable.extract(worksheet_name, i).each do |table_name, details|
            name = table_name.downcase
            table = Table.new(table_name, *details)
            @tables[name] = table
            @table_areas[name.to_sym] = table.all
            @table_data[name.to_sym] = table.data
          end
        end
      end
    end

    # Replace A$1:B2 with [A1, A2, B1, B2]
    @replace_ranges_with_array_literals_replacer ||= ReplaceRangesWithArrayLiteralsAst.new

    @table_areas.each do |name, reference|
      @table_areas[name] = @replace_ranges_with_array_literals_replacer.map(reference)
    end

    @table_data.each do |name, reference|
      @table_data[name] = @replace_ranges_with_array_literals_replacer.map(reference)
    end
    
  end

  def check_all_functions_implemented
    functions_that_are_removed_during_compilation = [:INDIRECT, :OFFSET, :ROW, :COLUMN, :TRANSPOSE]
    functions_used = CachingFormulaParser.instance.functions_used.keys
    functions_used.delete_if do |f|
      MapFormulaeToRuby::FUNCTIONS[f]
    end
    functions_that_are_removed_during_compilation.each do |f|
      functions_used.delete(f)
    end

    unless functions_used.empty?
      
      log.fatal "The following functions have not been implemented in excel_to_code #{ExcelToCode.version}:"

      functions_used.each do |f|
        log.fatal f.to_s
      end 
      
      log.fatal "Check for a new version of excel_to_code at https://github.com/tamc/excel_to_code"
      log.fatal "Or follow the instractions at https://github.com/tamc/excel_to_code/blob/master/doc/How_to_add_a_missing_function.md to implement the function yourself"
      exit
    end
  end
  
  # This makes sure that cells_to_keep includes named_references_to_keep
  def transfer_named_references_to_keep_into_cells_to_keep
    log.info "Transfering named references to keep into cells to keep"
    return unless @named_references_to_keep
    if @named_references_to_keep == :all
      @named_references_to_keep = @named_references.keys + @table_areas.keys 
      # If the user has specified named_references_to_keep == :all, but there are none, fall back
      if @named_references_to_keep.empty?
        log.warn "named_references_to_keep == :all, but no named references found"
        return
      end
    end
    @cells_to_keep ||= {}
    @named_references_to_keep.each do |name|
      ref = @named_references[name] || @table_areas[name]
      if ref
        add_ref_to_hash(ref, @cells_to_keep)
      else
        log.warn "Named reference \"#{name}\" not found"
      end
    end
  end

  # This makes sure that there are cell setter methods for any named references that can be set
  def transfer_named_references_that_can_be_set_at_runtime_into_cells_that_can_be_set_at_runtime
    log.info "Making sure there are setter methods for named references that can be set"
    return unless @named_references_that_can_be_set_at_runtime
    return if @named_references_that_can_be_set_at_runtime == :where_possible # in this case will be done in #work_out_which_named_references_can_be_set_at_runtime
    @cells_that_can_be_set_at_runtime ||= {}
    @named_references_that_can_be_set_at_runtime.each do |name|
      ref = @named_references[name] || @table_areas[name]
      if ref
        add_ref_to_hash(ref, @cells_that_can_be_set_at_runtime)
      else
        log.warn "Named reference #{name} not found"
      end
    end
  end

  # The reference passed may be a sheet reference or an area reference
  # in which case we need to expand out the ref so that the hash contains
  # one reference per cell
  def add_ref_to_hash(ref, hash)
    ref = ref.dup
    if ref.first == :sheet_reference
      sheet = ref[1]
      cell = Reference.for(ref[2][1]).unfix.to_sym
      hash[sheet] ||= []
      return if hash[sheet] == :all
      hash[sheet] << cell.to_sym unless hash[sheet].include?(cell.to_sym)
    elsif ref.first == :array
      ref.shift
      ref.each do |row|
        row = row.dup
        row.shift
        row.each do |cell|
          add_ref_to_hash(cell, hash)
        end
      end
    else
      log.error "Weird reference in named reference #{ref}"
    end
  end
  
  # Excel can include references to strings rather than the strings
  # themselves. This harmonises so the strings themselves are always
  # used.
  def rewrite_values_to_remove_shared_strings
    log.info "Rewriting values"
    r = ReplaceSharedStringAst.new(@shared_strings)
    @values.each do |ref, ast|
      r.map(ast)
    end
  end
  
  # In Excel we can have references like A:Z and 5:20 which mean all cells in columns 
  # A to Z and all cells in rows 5 to 20 respectively. This function translates these
  # into more conventional references (e.g., A5:Z20) based on the maximum area that 
  # has been used on a worksheet
  def rewrite_row_and_column_references
    log.info "Rewriting row and column references"
    # FIXME: Refactor
    dimension_objects = {}
    @worksheets_dimensions.map do |sheet_name, dimension| 
      dimension_objects[sheet_name] = WorksheetDimension.new(dimension) 
    end
    mapper = MapColumnAndRowRangeAst.new(nil, dimension_objects)

    @formulae_simple.each do |ref, ast|
      mapper.default_worksheet_name = ref.first
      mapper.map(ast)
    end

    @formulae_shared.each do |ref, ast|
      mapper.default_worksheet_name = ref.first
      mapper.map(ast.last)
    end

    @formulae_array.each do |ref, ast|
      mapper.default_worksheet_name = ref.first
      mapper.map(ast.last)
    end
    # FIXME: Could we now nil off the dimensions? Or do we need for indirects?
  end
  
  # Excel can share formula definitions across cells. This function unshares
  # them so every cell has its own definition
  def rewrite_shared_formulae_into_normal_formulae
    log.info "Rewriting shared formulae"
    @formulae_shared = RewriteSharedFormulae.rewrite( @formulae_shared, @formulae_shared_targets)
    @shared_formulae_targets = :no_longer_needed # Allow the targets to be garbage collected.
  end

  # Excel has the concept of array formulae: formulae whose answer spans
  # many cells. They are awkward. We try and replace them with conventional
  # formulae here.
  def rewrite_array_formulae
    log.info "Expanding #{@formulae_array.size} array formulae"
    # FIMXE: Refactor this

    named_reference_replacer = ReplaceNamedReferencesAst.new(@named_references, nil, @table_data)
    table_reference_replacer = ReplaceTableReferenceAst.new(@tables)
    @replace_ranges_with_array_literals_replacer ||= ReplaceRangesWithArrayLiteralsAst.new
    expand_array_formulae_replacer = AstExpandArrayFormulae.new
    simplify_arithmetic_replacer ||= SimplifyArithmeticAst.new
    @shared_string_replacer ||= ReplaceSharedStringAst.new(@shared_strings)
    transpose_function_replacer = ReplaceTransposeFunction.new

    # FIXME: THIS IS THE MOST HORRIFIC BODGE. I HATE IT.
    emergency_indirect_replacement_bodge = EmergencyArrayFormulaReplaceIndirectBodge.new
    emergency_indirect_replacement_bodge.references = @values
    emergency_indirect_replacement_bodge.tables = @tables
    emergency_indirect_replacement_bodge.named_references = @named_references
    
    @formulae_array.each do |ref, details|
      begin
        @shared_string_replacer.map(details.last)
        emergency_indirect_replacement_bodge.current_sheet_name = ref.first
        emergency_indirect_replacement_bodge.referring_cell = ref.last
        emergency_indirect_replacement_bodge.replace(details.last)

        named_reference_replacer.default_sheet_name = ref.first
        named_reference_replacer.map(details.last)
        table_reference_replacer.worksheet = ref.first
        table_reference_replacer.referring_cell = ref.last
        table_reference_replacer.map(details.last)
        @replace_ranges_with_array_literals_replacer.map(details.last)
        transpose_function_replacer.map(details.last)
        simplify_arithmetic_replacer.map(details.last)
        # FIXME: Seem to need to do this twice, second time to eliminate brackets?!
        simplify_arithmetic_replacer.map(details.last)
        expand_array_formulae_replacer.map(details.last)
      rescue  Exception => e
        log.fatal "Exception when expanding array formulae #{ref}: #{details}"
        raise
      end
    end

    log.info "Rewriting array formulae into conventional formulae"
    @formulae_array = RewriteArrayFormulae.rewrite(@formulae_array)
  end

  # At the end of this function we are left with a single @formulae hash
  # that contains every cell in the workbook, whatever its original format.
  def combine_formulae_types
    log.info "Combining formulae types"

    @formulae = required_references
    # We dup this to avoid the values being replaced when manipulating formulae
    @values.each do |ref, value|
      @formulae[ref] = value.dup
    end
    @formulae.merge! @formulae_shared
    @formulae.merge! @formulae_array
    @formulae.merge! @formulae_simple

    log.info "Sheet contains #{@formulae.size} cells"
  end
  
  # Turns aritmetic with many arguments (1+2+3+4) into arithmetic with only
  # two arguments (((1+2)+3)+4), taking into account operator precedence.
  def simplify_arithmetic
    log.info "Simplifying arithmetic"
    simplify_arithmetic_replacer ||= SimplifyArithmeticAst.new
    @formulae.each do |ref, ast|
      simplify_arithmetic_replacer.map(ast)
    end
  end
  
  # This ensures that all gettable and settable values appear in the output
  # even if they are blank in the underlying excel
  def required_references
    log.info "Checking required references"
    required_refs = {}

    # Need to add blank for any settable cells that aren't defined
    if @cells_that_can_be_set_at_runtime && @cells_that_can_be_set_at_runtime != :named_references_only
      @cells_that_can_be_set_at_runtime.each do |worksheet, refs|
        next if refs == :all
        refs.each do |ref|
          required_refs[[worksheet, ref]] = [:blank]
        end
      end
    end

    # Need to add blanks for any cells the user want's, but aren't defined
    if @cells_to_keep
      @cells_to_keep.each do |worksheet, refs|
        next if refs == :all
        refs.each do |ref|
          required_refs[[worksheet, ref]] = [:blank]
        end
      end
    end

    # In some situations also need to add the named references 
    if @named_references_to_keep
      @named_references_to_keep.each do |name|
        ref = @named_references[name] || @table_areas[name]
        if ref.first == :sheet_reference
          s = ref[1]
          c = Reference.for(ref[2][1]).unfix.to_sym
          required_refs[[s, c]] = [:blank]
        elsif ref.first == :array
          ref.each do |row|
            next unless row.is_a?(Array)
            row.each do |cell|
              next unless cell.is_a?(Array)
              if cell.first == :sheet_reference
                s = cell[1]
                c = Reference.for(cell[2][1]).unfix.to_sym
                required_refs[[s, c]] = [:blank]
              else
                log.warn "Named reference '#{name}' refers to cells that can't be interpreted"
              end
            end
          end
        else
          log.warn "Named reference '#{name}' refers to cells that can't be interpreted"
        end
      end
    end
    required_refs
  end

  # This just checks which named references refer to cells that we have already declared as settable
  def work_out_which_named_references_can_be_set_at_runtime
    log.info "Working out which named references can be set at runtime"
    return unless @named_references_that_can_be_set_at_runtime
    return unless @named_references_that_can_be_set_at_runtime == :where_possible
    cells_that_can_be_set = @cells_that_can_be_set_at_runtime
    cells_that_can_be_set = cells_with_settable_values if cells_that_can_be_set == :named_references_only
    cells_that_can_be_set_due_to_named_reference = Hash.new { |h,k| h[k] = Array.new  }
    @named_references_that_can_be_set_at_runtime = []
    all_named_references = @named_references
    # FIXME can this be refactored with #add_ref_to_hash
    @named_references_to_keep.each do |name|
      ref = all_named_references[name]
      unless ref
        log.warn "Named reference to keep #{name} not found in spreadsheet"
        next
      end
      if ref.first == :sheet_reference
        sheet = ref[1]
        cell = Reference.for(ref[2][1]).unfix.to_sym
        s = cells_that_can_be_set[sheet]
        if s && ( s == :all || s.include?(cell) )
          @named_references_that_can_be_set_at_runtime << name 
          cells_that_can_be_set_due_to_named_reference[sheet] << cell.to_sym
          cells_that_can_be_set_due_to_named_reference[sheet].uniq!
        end
        #FIXME: Is this righ?
      elsif ref.first.is_a?(Array)
        ref = ref.first
        settable = ref.all? do |r|
          sheet = r[1]
          cell = r[2][1].gsub('$','')
          s = cells_that_can_be_set[sheet]
          s && s.include?(cell)
        end
        if settable
          @named_references_that_can_be_set_at_runtime << name 
          ref.each do |r| 
            sheet = r[1]
            cell = r[2][1].gsub('$','')
            cells_that_can_be_set_due_to_named_reference[sheet] << cell.to_sym
            cells_that_can_be_set_due_to_named_reference[sheet].uniq!
          end
        end
      end
    end
    if @cells_that_can_be_set_at_runtime == :named_references_only
      @cells_that_can_be_set_at_runtime = cells_that_can_be_set_due_to_named_reference
    end
  end

  # FIXME: Feels like a kludge
  # This works out which named references should appear in the generated code
  def filter_named_references
    log.info "Filtering named references to keep"
    @named_references_to_keep ||= []
    @named_references_that_can_be_set_at_runtime ||= []

    @named_references.each do |name, ref|
      if named_references_to_keep.include?(name) || named_references_that_can_be_set_at_runtime.include?(name)
        # FIXME: Refactor the c_name_for to closer to the writing?
        @named_references_to_keep << name
      end
    end

  end
    
  def simplify(cells = @formulae)
    log.info "Simplifying cells"

    @shared_string_replacer ||= ReplaceSharedStringAst.new(@shared_strings)
    @replace_arithmetic_on_ranges_replacer ||= ReplaceArithmeticOnRangesAst.new
    @wrap_formulae_that_return_arrays_replacer ||= WrapFormulaeThatReturnArraysAndAReNotInArraysAst.new
    @named_reference_replacer ||= ReplaceNamedReferencesAst.new(@named_references, nil, @table_data) 
    @table_reference_replacer ||= ReplaceTableReferenceAst.new(@tables)
    @replace_ranges_with_array_literals_replacer ||= ReplaceRangesWithArrayLiteralsAst.new
    @replace_arrays_with_single_cells_replacer ||= ReplaceArraysWithSingleCellsAst.new
    @replace_string_joins_on_ranges_replacer ||= ReplaceStringJoinOnRangesAST.new
    @sheetless_cell_reference_replacer ||= RewriteCellReferencesToIncludeSheetAst.new
    @replace_references_to_blanks_with_zeros ||= ReplaceReferencesToBlanksWithZeros.new(@formulae, nil, inline_ast_decision)
    @fix_subtotal_of_subtotals ||= FixSubtotalOfSubtotals.new(@formulae)
    # FIXME: Bodge to put it here as well, but seems to be required
    column_and_row_function_replacement = ReplaceColumnAndRowFunctionsAST.new

    #require 'pry'; binding.pry

    cells.each do |ref, ast|
      begin
        @sheetless_cell_reference_replacer.worksheet = ref.first
        @sheetless_cell_reference_replacer.map(ast)
        @shared_string_replacer.map(ast)
        @named_reference_replacer.default_sheet_name = ref.first
        @named_reference_replacer.map(ast)
        @table_reference_replacer.worksheet = ref.first
        @table_reference_replacer.referring_cell = ref.last
        @table_reference_replacer.map(ast)
        @replace_ranges_with_array_literals_replacer.map(ast)
        @replace_arrays_with_single_cells_replacer.ref = ref
        a = @replace_arrays_with_single_cells_replacer.map(ast)
        if @replace_arrays_with_single_cells_replacer.need_to_replace
          cells[ref] = @formulae[ref] = a
        end
        @replace_arithmetic_on_ranges_replacer.map(ast)
        @replace_string_joins_on_ranges_replacer.map(ast)
        @wrap_formulae_that_return_arrays_replacer.map(ast)
        column_and_row_function_replacement.current_reference = ref.last
        column_and_row_function_replacement.replace(ast)
        @replace_references_to_blanks_with_zeros.current_sheet_name = ref.first
        @replace_references_to_blanks_with_zeros.map(ast)
        @fix_subtotal_of_subtotals.map(ast)
      rescue  Exception => e
        log.fatal "Exception when simplifying #{ref}: #{ast}"
        raise
      end
    end
  end

  # These types of cells don't conatain formulae and can therefore be skipped
  VALUE_TYPE = {:number => true, :string => true, :blank => true, :null => true, :error => true, :boolean_true => true, :boolean_false => true}

  def must_keep?(ref)
    must_keep_in_sheet = @cells_that_can_be_set_at_runtime[ref.first]
    return false unless must_keep_in_sheet
    if must_keep_in_sheet == :all
      # Only keep cells that actually exist
      return true if @values[ref]
      return false
    end
    must_keep_in_sheet.include?(ref.last)
  end
  
  def inline_ast_decision  
    @inline_ast_decision ||= lambda do |sheet, cell, references|
      ref = [sheet,cell]
      if must_keep?(ref)
        false
      else
        ast = references[ref]
        if ast
          case ast.first
          when :number, :string; true
          when :blank, :null; true
          when :error; true
          when :boolean_true, :boolean_false; true
          else
            false
          end
        else
          true # Always inline blanks
        end
      end
    end
  end

  def replace_formulae_with_their_results
    number_of_passes = 0

    @cells_with_formulae = @formulae.dup
    @cells_with_formulae.each do |ref, ast|
      @cells_with_formulae.delete(ref) if VALUE_TYPE[ast[0]]
    end

    # Set up for replacing references to cells with the cell
    inline_replacer = InlineFormulaeAst.new
    inline_replacer.references = @formulae
    inline_replacer.inline_ast = inline_ast_decision

    value_replacer = MapFormulaeToValues.new
    value_replacer.original_excel_filename = excel_file

    # There is no support for INDIRECT or OFFSET in the ruby or c runtime
    # However, in many cases it isn't needed, because we can work
    # out the value of the indirect or OFFSET at compile time and eliminate it
    # First of all we replace any indirects where their values can be calculated at compile time with those
    # calculated values (e.g., INDIRECT("A"&1) can be turned into A1 and OFFSET(A1,1,1,2,2) can be turned into B2:C3)
    indirect_replacement = ReplaceIndirectsWithReferencesAst.new
    column_and_row_function_replacement = ReplaceColumnAndRowFunctionsAST.new
    offset_replacement = ReplaceOffsetsWithReferencesAst.new

    begin 
      number_of_passes += 1
      log.info "Starting pass #{number_of_passes} on #{@cells_with_formulae.size} cells"

      replacements_made_in_the_last_pass = 0
      inline_replacer.count_replaced = 0
      value_replacer.replacements_made_in_the_last_pass = 0
      column_and_row_function_replacement.count_replaced = 0
      offset_replacement.count_replaced = 0
      indirect_replacement.count_replaced = 0
      references_that_need_updating = {}

      @cells_with_formulae.each do |ref, ast|
        begin
          column_and_row_function_replacement.current_reference = ref.last
          if column_and_row_function_replacement.replace(ast)
            references_that_need_updating[ref] = ast
          end
          if offset_replacement.replace(ast)
            references_that_need_updating[ref] = ast
          end
          # FIXME: Shouldn't need to wrap ref.fist in an array
          inline_replacer.current_sheet_name = [ref.first]
          inline_replacer.map(ast)
          # If a formula references a cell containing a value, the reference is replaced with the value (e.g., if A1 := 2 and A2 := A1 + 1 then becomes: A2 := 2 + 1)
          #require 'pry'; binding.pry if ref == [:"Outputs - Summary table", :E77]
          value_replacer.map(ast)
          if indirect_replacement.replace(ast)
            references_that_need_updating[ref] = ast
          end
          @cells_with_formulae.delete(ref) if VALUE_TYPE[ast[0]]
        rescue  Exception => e
          log.fatal "Exception when replacing formulae with results in #{ref}: #{ast}"
          raise
        end
      end
      

      @named_references.each do |ref, ast|
        inline_replacer.current_sheet_name = ref.is_a?(Array) ? [ref.first] : []
        inline_replacer.map(ast)
      end

      simplify(references_that_need_updating)

      replacements_made_in_the_last_pass += inline_replacer.count_replaced
      replacements_made_in_the_last_pass += value_replacer.replacements_made_in_the_last_pass
      replacements_made_in_the_last_pass += column_and_row_function_replacement.count_replaced
      replacements_made_in_the_last_pass += offset_replacement.count_replaced
      replacements_made_in_the_last_pass += indirect_replacement.count_replaced

      log.info "Pass #{number_of_passes}: Made #{replacements_made_in_the_last_pass} replacements"
    end while replacements_made_in_the_last_pass > 0 && number_of_passes < 20
  end

  
  
  # If 'cells to keep' are specified, then other cells are removed, unless
  # they are required to calculate the value of a cell in 'cells to keep'.
  def remove_any_cells_not_needed_for_outputs
    log.info "Removing cells not needed for outputs"

    # If 'cells to keep' isn't specified, then ALL cells are kept
    return unless cells_to_keep && !cells_to_keep.empty?
    
    # Work out what cells the cells in 'cells to keep' need 
    # in order to be able to calculate their values
    identifier = IdentifyDependencies.new
    identifier.references = @formulae
    cells_to_keep.each do |sheet_to_keep,cells_to_keep|
      if cells_to_keep == :all
        identifier.add_depedencies_for(sheet_to_keep)
      elsif cells_to_keep.is_a?(Array)
        cells_to_keep.each do |cell|
          identifier.add_depedencies_for(sheet_to_keep,cell)
        end
      end
    end
        
    # On top of that, we don't want to remove any cells
    # that have been specified as 'settable'
    worksheets do |name,xml_filename|
      s = @cells_that_can_be_set_at_runtime[name]
      next unless s
      if s == :all
        identifier.add_depedencies_for(name)
      else 
        s.each do |ref|
          identifier.add_depedencies_for(name,ref)
        end
      end
    end
    
    # Now we actually go ahead and remove the cells
    r = RemoveCells.new
    r.cells_to_keep = identifier.dependencies
    r.rewrite(@formulae)
    # Must remove the values as well, to avoid any tests being generated for cells that don't exist
    r.rewrite(@values)
    r.rewrite(@cells_with_formulae)
  end
  
  # If a cell is only referenced from one other cell, then it is inlined into that other cell
  # e.g., A1 := B3+B6 ; B1 := A1 + B3 becomes: B1 := (B3 + B6) + B3. A1 is removed.
  def inline_formulae_that_are_only_used_once
    log.info "Inlining formulae"

    # First step is to calculate how many times each cell is referenced by another cell
    counter = CountFormulaReferences.new
    count = counter.count(@formulae)
    
    # This takes the decision:
    # 1. If a cell is in the list of cells to keep, then it is never inlined
    # 2. Otherwise, it is inlined if only one other cell refers to it.
    inline_ast_decision = lambda do |sheet,cell,references|
      references_to_keep = @cells_that_can_be_set_at_runtime[sheet]
      if references_to_keep && (references_to_keep == :all || references_to_keep.include?(cell))
        false
      else
        count[[sheet,cell]] == 1 # i.e., inline if used only once
      end
    end
    
    r = InlineFormulaeAst.new
    r.references = @formulae
    r.inline_ast = inline_ast_decision
    @cells_with_formulae.each do |ref, ast|
      begin
        r.current_sheet_name = [ref.first]
        r.map(ast)
      rescue  Exception => e
        log.fatal "Exception when inlining formulae only used once in #{ref}: #{ast}"
        raise
      end
    end
  end
  
  # This comes up with a list of references to test, in the form of a file called 'References to test'.
  # It is structured to contain one reference per row:
  # worksheet_c_name \t ref \t value_ast
  # These will be sorted so that later refs depend on earlier refs. This should mean that the first test that 
  # fails will be the root cause of the problem
  def create_sorted_references_to_test
    log.info "Creating references to test"

    references_to_test = {}

    # First get the list of references we should test
    @values.each do |ref, value|
      if !cells_to_keep || 
          cells_to_keep.empty? || 
          (cells_to_keep[ref.first] && (
            cells_to_keep[ref.first] == :all ||
            cells_to_keep[ref.first].include?(ref.last)
          ))
        references_to_test[ref] = value
      end
    end

    # Now work out dependency tree
    sorted_references = @formulae.keys #SortIntoCalculationOrder.new.sort(@formulae)

    @references_to_test_array = []
    sorted_references.each do |ref|
      next unless references_to_test.include?(ref)
      @references_to_test_array << [ref, @values[ref]]
    end
    # FIXME: CNAMES
  end


  # This looks for repeated formula parts, and separates them out. It is the opposite of inlining:
  # e.g., A1 := (B1 + B3) + B10; A2 := (B1 + B3) + 3 gets transformed to: Common1 := B1 + B3 ; A1 := Common1 + B10 ; A2 := Common1 + 3
  def separate_formulae_elements
    log.info "Looking for repeated bits of formulae"
    
    
    identifier = IdentifyRepeatedFormulaElements.new
    repeated_elements = identifier.count(@cells_with_formulae)
    
    # We apply a threshold that something needs to be used twice for us to bother separating it out. 
    # FIXME: This threshold is arbitrary
    repeated_elements.delete_if do |element,count|
      count < 2
    end
    
    # Translate the repeated elements into a code of the form [:cell, "common#{1}"]
    index = 0
    repeated_element_ast = {}
    repeated_elements.each do |ast, count|
      repeated_element_ast[ast.dup] = [:cell, "common#{index}"]
      index +=1 
    end

    r = ReplaceCommonElementsInFormulae.new
    r.replace(@cells_with_formulae, repeated_element_ast)
    common_elements_used = r.common_elements_used

    repeated_element_ast.delete_if do |repeated_ast, common_ast|
      common_elements_used[common_ast] == 0
    end

    # FIXME: Is this best? Seems to work
    repeated_element_ast.each do |repeated_ast, common_ast|
      @formulae[["", common_ast[1]]] = repeated_ast
    end

  end
  
  # This puts back in an optimisation that excel carries out by making sure that
  # two copies of the same value actually refer to the same underlying spot in memory
  def replace_values_with_constants
    log.info "Replacing values with constants"
    
    # First do it in the formulae
    r = MapValuesToConstants.new
    @formulae.each do |ref, ast|
      r.map(ast)
    end

    @named_references.each do |ref, ast|
      r.map(ast)
    end

    @constants = r.constants.invert
  end
  
  # If nothing has been specified in named_references_that_can_be_set_at_runtime 
  # or in cells_that_can_be_set_at_runtime, then we assume that
  # all value cells should be settable if they are referenced by
  # any other forumla.
  def ensure_there_is_a_good_set_of_cells_that_can_be_set_at_runtime
    # By this stage, if named_references were set, then cells_that_can_be_set_at_runtime will 
    # have been set to match
    return unless @cells_that_can_be_set_at_runtime.empty?
    @cells_that_can_be_set_at_runtime = cells_with_settable_values
  end

  # Returns a list of cells that are:
  # 1. Simple values (e.g., a string or a number)
  # 2. That are referenced in other formulae
  # ... these are likely to be cells that the user will want to use as inputs
  # to their calculation.
  def cells_with_settable_values
    log.info "Generating a good set of cells that should be settable"

    counter = CountFormulaReferences.new
    count = counter.count(@formulae)
    settable_cells = {}
    settable_types = [:blank,:number,:null,:string,:shared_string,:constant,:percentage,:error,:boolean_true,:boolean_false]

    count.each do |ref,count|
      next unless count >= 1 # No point making a cell that isn't reference settable
      ast = @formulae[ref]
      next unless ast # Sometimes empty cells are referenced. 
      next unless settable_types.include?(ast.first)
      settable_cells[ref.first] ||= []
      settable_cells[ref.first] << ref.last.upcase
    end
    return settable_cells
  end
  
  # UTILITY FUNCTIONS

  def settable
    settable_refs = @cells_that_can_be_set_at_runtime
    if settable_refs
      lambda { |ref| 
        sheet = ref.first
        cell = ref.last
        if settable_refs[sheet]
          if settable_refs[sheet] == :all || settable_refs[sheet].include?(cell.upcase)
            true
          else
            false
          end
        else
          false
        end
      }
    else
      lambda { |ref| false }
    end
  end
  
  def gettable
    if @cells_to_keep
      gettable_refs = @cells_to_keep
      lambda { |ref| 
        sheet = ref.first
        cell = ref.last
        if gettable_refs[sheet]
          if gettable_refs[sheet] == :all || gettable_refs[sheet].include?(cell.upcase)
            true
          else
            false
          end
        else
          false
        end
      }
    else
      lambda { |ref| true }
    end
  end
    
  def c_name_for_worksheet_name(name)
    @worksheet_c_names[name.to_s]
  end
    
  def worksheets
    @worksheet_xmls.each do |name, filename|
      yield name, filename
    end
  end
  
  def xml(*args, &block)
    args.flatten!
    filename = File.join(xml_directory,'xl',*args)
    if File.exists?(filename)
      f = File.open(filename,'r')
    else
      log.warn("#{filename} does not exist in xml(#{args.inspect}), using blank instead")
      f = StringIO.new
    end
    if block
      yield f
      f.close if f.respond_to?(:close)
    else
      f
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
  
  def ruby_module_name
    @ruby_module_name = output_name.sub(/^[a-z\d]*/) { $&.capitalize }
    @ruby_module_name = @ruby_module_name.gsub(/(?:_|(\/))([a-z\d]*)/i) { "#{$1}#{$2.capitalize}" }.gsub('/', '::')
    @ruby_module_name
  end

  def cleanup
    log.info "Cleaning up"
    FileUtils.remove_entry(self.xml_directory) if @delete_xml_directory_at_end
  end

end
