# coding: utf-8
require 'fileutils'
require 'logger'
require_relative '../excel_to_code'

# Used to throw normally fatal errors
class ExcelToCodeException < Exception; end
class VersionedFileNotFoundException < Exception; end
class XMLFileNotFoundException < Exception; end

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

  # Optional attribute. The intermediate workings will be stored here.
  # If not specified, will be '#{excel_file_name}/intermediate'
  attr_accessor :intermediate_directory
  
  # Optional attribute. Specifies which cells have setters created in the c code so their values can be altered at runtime.
  # It is a hash. The keys are the sheet names. The values are either the symbol :all to specify that all cells on that sheet 
  # should be setable, or an array of cell names on that sheet that should be settable (e.g., A1)
  attr_accessor :cells_that_can_be_set_at_runtime
  
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
  
  # Optional attribute. Boolean. Not relevant to all types of code output
  #   * true - the generated c code is compiled
  #   * false - the generated c code is not compiled (default, unless actuall_run_tests is specified as true)
  attr_accessor :actually_compile_code

  # Optional attribute. Boolean. 
  #   * true - the generated tests are run
  #   * false - the generated tests are not run
  attr_accessor :actually_run_tests
  
  # Optional attribute. Boolean.
  #   * true - the intermediate files are not written to disk (requires a lot of memory)
  #   * false - the intermediate files are written to disk (default, easier to debug)
  attr_accessor :run_in_memory
  
  # This is the log file, if set it needs to respond to the same methods as the standard logger library
  attr_accessor :log

  # Optional attribute. Boolean. Default true.
  #   * true - empty cells and zeros are treated as being equivalent in tests. Numbers greater then 1 are only expected to match with assert_in_epsilon, numbers less than 1 are only expected to match with assert_in_delta
  #   * false - empty cells and zeros are treated as being different in tests. Numbers must match to full accuracy.
  attr_accessor :sloppy_tests
  
  def set_defaults
    raise ExcelToCodeException.new("No excel file has been specified") unless excel_file
    
    self.output_directory ||= File.join(File.dirname(excel_file),File.basename(excel_file,".*"),language)
    self.xml_directory ||= File.join(File.dirname(excel_file),File.basename(excel_file,".*"),'xml')
    self.intermediate_directory ||= File.join(File.dirname(excel_file),File.basename(excel_file,".*"),'intermediate')
    
    self.output_name ||= "Excelspreadsheet"
    
    self.cells_that_can_be_set_at_runtime ||= {}
    
    # Make sure that all the cell names are downcase and don't have any $ in them
    cells_that_can_be_set_at_runtime.keys.each do |sheet|
      next unless cells_that_can_be_set_at_runtime[sheet].is_a?(Array)
      cells_that_can_be_set_at_runtime[sheet] = cells_that_can_be_set_at_runtime[sheet].map { |reference| reference.gsub('$','').upcase }
    end

    # Make sure that all the cell names are downcase and don't have any $ in them
    if cells_to_keep
      cells_to_keep.keys.each do |sheet|
        next unless cells_to_keep[sheet].is_a?(Array)
        cells_to_keep[sheet] = cells_to_keep[sheet].map { |reference| reference.gsub('$','').upcase }
      end
    end  
    
    # Make sure the relevant directories exist
    self.excel_file = File.expand_path(excel_file)
    self.output_directory = File.expand_path(output_directory)
    
    # Set up our log file
    self.log ||= Logger.new(STDOUT)

    # By default, tests allow empty cells and zeros to be treated as equivalent, and numbers only have to match to a 0.001 epsilon (if expected>1) or 0.001 delta (if expected<1)
    self.sloppy_tests ||= true
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
    merge_table_files
    
    # These perform some translations to simplify the excel
    # Including:
    # * Turning row and column references (e.g., A:A) to areas, based on the size of the worksheet
    # * Turning range references (e.g., A1:B2) into array litterals (e.g., {A1,B1;A2,B2})
    # * Turning shared formulae into a series of conventional formulae
    # * Turning array formulae into a series of conventional formulae
    # * Mergining all the different types of formulae and values into a single file
    rewrite_worksheets
    
    # In case this hasn't been set by the user
    if cells_that_can_be_set_at_runtime.empty?
      log.info "Creating a good set of cells that should be settable"
      create_a_good_set_of_cells_that_should_be_settable_at_runtime
    end
    
    # These perform a series of transformations to the information
    # with the intent of removing any redundant calculations
    # that are in the excel.
    simplify_worksheets # Replacing shared strings and named references with their actual values, tidying arithmetic
    replace_formulae_with_their_results
    remove_any_cells_not_needed_for_outputs
    inline_formulae_that_are_only_used_once
    separate_formulae_elements
    replace_values_with_constants
        
    # This actually creates the code (implemented in subclasses)
    write_code
    
    # These compile and run the code version of the excel (implemented in subclasses)
    compile_code
    run_tests
    
    log.info "The generated code is available in #{File.join(output_directory)}"
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
  end
  
  # Excel keeps a list of worksheet names. To get the mapping between
  # human and computer name  correct we have to look in the workbook 
  # relationships files. We also need to mangle the name into something
  # that will work ok as a filesystem or program name
  def extract_worksheet_names
    extract ExtractWorksheetNames, 'workbook.xml', 'Worksheet names'
    extract ExtractRelationships, File.join('_rels','workbook.xml.rels'), 'Workbook relationships'
    rewrite RewriteWorksheetNames, 'Worksheet names', 'Workbook relationships', 'Worksheet names'
    rewrite MapSheetNamesToCNames, 'Worksheet names', 'Worksheet C names'
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
  # 
  # It then looks at the relationship file and extracts any tables
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
      
      extract_tables_for_worksheet(name,xml_filename)
    end
  end
  
  # To extract a table we need to look in the worksheet for table references
  # then we look in the relationships file for the filename that matches that
  # reference and contains the table data. Then we consolidate all the data
  # from individual table files into a single table file for the worksheet.
  def extract_tables_for_worksheet(name, xml_filename)
    extract ExtractWorksheetTableRelationships, xml_filename, [name, "Worksheet tables"]
    extract ExtractRelationships, File.join('worksheets','_rels',"#{File.basename(xml_filename)}.rels"), [name, 'Relationships']
    rewrite RewriteRelationshipIdToFilename, [name, "Worksheet tables"], [name, 'Relationships'], [name, "Worksheet tables"]
    table_filenames = input(name, "Worksheet tables")
    tables = intermediate(name, "Worksheet tables")
    table_extractor = ExtractTable.new(name)
    table_filenames.lines.each do |line|
      table_xml = xml(File.join('worksheets',line.strip))
      table_extractor.extract(table_xml, tables)
    end
    close(tables,table_filenames)
  end
  
  # Tables are like named references in that they can be referred to from
  # anywhere in the workbook. Therefore we consolidate all the tables from
  # all the worksheets into a central table file.
  def merge_table_files
    merged_table_file = intermediate("Workbook tables")
    worksheets do |name,xml_filename|
      log.info "Merging table files for #{name}"
      worksheet_table_file = input([name, "Worksheet tables"])
      worksheet_table_file.lines do |line|
        merged_table_file.puts line
      end
      close worksheet_table_file
    end
    close merged_table_file
  end
  
  def rewrite_worksheets
    worksheets do |name,xml_filename|
      log.info "Rewriting worksheet #{name}"
      rewrite_row_and_column_references(name,xml_filename)
      rewrite_shared_formulae(name,xml_filename)
      rewrite_array_formulae(name,xml_filename)
      combine_formulae_files(name,xml_filename)
    end
  end
  
  def rewrite_row_and_column_references(name,xml_filename)
    dimensions = input('Worksheet dimensions')
    
    r = RewriteWholeRowColumnReferencesToAreas.new
    r.worksheet_dimensions = dimensions
    r.sheet_name = name
    
    apply_rewrite r, [name, 'Formulae (simple)']
    apply_rewrite r, [name, 'Formulae (shared)']
    apply_rewrite r, [name, 'Formulae (array)']
    
    dimensions.close
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
    combiner.references_to_add_if_they_are_not_already_present = required_references(name)
    
    rewrite combiner, [name, 'Values'], [name, 'Formulae (shared)'], [name, 'Formulae (array)'], [name, 'Formulae (simple)'], [name, 'Formulae']
  end
  
  # This ensures that all gettable and settable values appear in the output
  # even if they are blank in the underlying excel
  def required_references(worksheet_name)
    required_refs = []
    if @cells_that_can_be_set_at_runtime && @cells_that_can_be_set_at_runtime[worksheet_name] && @cells_that_can_be_set_at_runtime[worksheet_name] != :all
      required_refs.concat(@cells_that_can_be_set_at_runtime[worksheet_name])
    end
    if @cells_to_keep && @cells_to_keep[worksheet_name] && @cells_to_keep[worksheet_name] != :all
      required_refs.concat(@cells_to_keep[worksheet_name])
    end
    required_refs
  end
    
  def simplify_worksheets
    worksheets do |name,xml_filename|
      replace ReplaceSharedStrings, [name, 'Values'], 'Shared strings', File.join(name, 'Values')
      
      replace SimplifyArithmetic,   [name, 'Formulae'], [name, 'Formulae']      
      replace ReplaceSharedStrings, [name, 'Formulae'], 'Shared strings', [name, 'Formulae']
      
      r = ReplaceNamedReferences.new
      r.sheet_name = name
      replace r, [name, 'Formulae'], 'Named references', [name, 'Formulae']

      r = ReplaceTableReferences.new
      r.sheet_name = name
      replace r, [name, 'Formulae'], "Workbook tables", [name, 'Formulae']
      
      replace ReplaceRangesWithArrayLiterals, [name, 'Formulae'],  [name, 'Formulae']
    end
  end
    
  # FIXME: This should work out how often it needs to operate, rather than having a hardwired 4
  def replace_formulae_with_their_results
    4.times do 
      replace_indirects
      replace_formulae_with_calculated_values
      replace_references_to_values_with_values
    end
  end
  
  # There is no support for INDIRECT in the ruby or c runtime
  # However, in many cases it isn't needed, because we can work
  # out the value of the indirect at compile time and eliminate it
  def replace_indirects
    worksheets do |name,xml_filename|
      log.info "Replacing indirects in #{name}"
      
      # First of all we replace any indirects where their values can be calculated at compile time with those
      # calculated values (e.g., INDIRECT("A"&1) can be turned into A1)
      replace ReplaceIndirectsWithReferences, [name, 'Formulae'],  [name, 'Formulae']
      
      # The result of the indirect might be a named reference, which we need to simplify
      r = ReplaceNamedReferences.new
      r.sheet_name = name
      replace r, [name, 'Formulae'], 'Named references', [name, 'Formulae']
      
      # The result of the indirect might be a table reference, which we need to simplify
      r = ReplaceTableReferences.new
      r.sheet_name = name
      replace r, [name, 'Formulae'], "Workbook tables", [name, 'Formulae']
      
      # The result of the indirect might be a range, which we need to simplify
      replace ReplaceRangesWithArrayLiterals, [name, 'Formulae'],  [name, 'Formulae']
      replace ReplaceArraysWithSingleCells, [name, 'Formulae'],  [name, 'Formulae']
    end
  end
  
  # If a formula's value can be calculated at compile time, it is replaced with its calculated value (e.g., 1+1 gets replaced with 2)
  def replace_formulae_with_calculated_values    
    worksheets do |name,xml_filename|
      replace ReplaceFormulaeWithCalculatedValues, [name, 'Formulae'],  [name, 'Formulae']
    end
  end

  # If a formula references a cell containing a value, the reference is replaced with the value (e.g., if A1 := 2 and A2 := A1 + 1 then becomes: A2 := 2 + 1)
  def replace_references_to_values_with_values
    references = all_formulae
    
    inline_ast_decision = lambda do |sheet,cell,references|
      references_to_keep = @cells_that_can_be_set_at_runtime[sheet]
      if references_to_keep && (references_to_keep == :all || references_to_keep.include?(cell))
        false
      else
        ast = references[sheet][cell]
        if ast
          if [:number,:string,:blank,:null,:error,:boolean_true,:boolean_false,:sheet_reference,:cell].include?(ast.first)
            true
          else
            false
          end
        else
          true # Always inline blanks
        end
      end
    end
    
    r = InlineFormulae.new
    r.references = references
    r.inline_ast = inline_ast_decision
    
    worksheets do |name,xml_filename|
      r.default_sheet_name = name
      replace r, [name, 'Formulae'],  [name, 'Formulae']
    end
  end
  
  # If 'cells to keep' are specified, then other cells are removed, unless
  # they are required to calculate the value of a cell in 'cells to keep'.
  def remove_any_cells_not_needed_for_outputs

    # If 'cells to keep' isn't specified, then ALL cells are kept
    return unless cells_to_keep && !cells_to_keep.empty?
    
    # Work out what cells the cells in 'cells to keep' need 
    # in order to be able to calculate their values
    identifier = IdentifyDependencies.new
    identifier.references = all_formulae
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
    worksheets do |name,xml_filename|
      r = RemoveCells.new
      r.cells_to_keep = identifier.dependencies[name]
      rewrite r, [name, 'Formulae'],  [name, 'Formulae']
      rewrite r, [name, 'Values'],  [name, 'Values'] # Must remove the values as well, to avoid any tests being generated for cells that don't exist
    end
  end
  
  # If a cell is only referenced from one other cell, then it is inlined into that other cell
  # e.g., A1 := B3+B6 ; B1 := A1 + B3 becomes: B1 := (B3 + B6) + B3. A1 is removed.
  def inline_formulae_that_are_only_used_once
    references = all_formulae
    
    # First step is to calculate how many times each cell is referenced by another cell
    counter = CountFormulaReferences.new
    count = counter.count(references)
    
    # This takes the decision:
    # 1. If a cell is in the list of cells to keep, then it is never inlined
    # 2. Otherwise, it is inlined if only one other cell refers to it.
    inline_ast_decision = lambda do |sheet,cell,references|
      references_to_keep = @cells_that_can_be_set_at_runtime[sheet]
      if references_to_keep && (references_to_keep == :all || references_to_keep.include?(cell))
        false
      else
        count[sheet][cell] == 1
      end
    end
    
    r = InlineFormulae.new
    r.references = references
    r.inline_ast = inline_ast_decision
    
    worksheets do |name,xml_filename|
      r.default_sheet_name = name
      replace r, [name, 'Formulae'],  [name, 'Formulae']
    end
    
    # We need to do this again, to get rid of the cells that we have just inlined
    # FIXME: This could be done more efficiently, given we know which cells were removed
    remove_any_cells_not_needed_for_outputs
  end
  
  # This looks for repeated formula parts, and separates them out. It is the opposite of inlining:
  # e.g., A1 := (B1 + B3) + B10; A2 := (B1 + B3) + 3 gets transformed to: Common1 := B1 + B3 ; A1 := Common1 + B10 ; A2 := Common1 + 3
  def separate_formulae_elements
    
    replace_all_simple_references_with_sheet_references # So we can be sure which references are repeating and which references are distinct
    
    references = all_formulae
    identifier = IdentifyRepeatedFormulaElements.new
    repeated_elements = identifier.count(references)
    
    # We apply a threshold that something needs to be used twice for us to bother separating it out. 
    # FIXME: This threshold is arbitrary
    repeated_elements.delete_if do |element,count|
      count < 2
    end
    
    # Dump our selected common elements into a separate file of formulae
    o = intermediate('Common elements')
    i = 0
    repeated_elements.each do |element,count|
      o.puts "common#{i}\t#{element}"
      i = i + 1
    end
    close(o)
    
    # Replace common elements in formulae with references to otherw
    worksheets do |name,xml_filename|
      replace ReplaceCommonElementsInFormulae, [name, 'Formulae'], "Common elements", [name, 'Formulae']
    end
    # FIXME: This means that some common elements won't ever be called, becuase they are replaced by a longer common element
    # Should the common elements be merged first?
  end

  # We add the sheet name to all references, so that we can then look for common elements accross worksheets
  # e.g., A1 := A2 gets transformed to A1 := Sheet1!A2  
  def replace_all_simple_references_with_sheet_references
    r = RewriteCellReferencesToIncludeSheet.new
    worksheets do |name,xml_filename|
      r.worksheet = name
      rewrite r, [name, 'Formulae'],  [name, 'Formulae']
    end
  end  
  
  # This puts back in an optimisation that excel carries out by making sure that
  # two copies of the same value actually refer to the same underlying spot in memory
  def replace_values_with_constants
    
    # First do it in the formulae
    r = ReplaceValuesWithConstants.new
    worksheets do |name,xml_filename|
      replace r, [name, 'Formulae'],  [name, 'Formulae']
    end
    
    # Then do it in the common elements
    replace r, "Common elements", "Common elements"
    
    # Then write out the constants
    output = intermediate("Constants")
    # FIXME: This looks bad!
    r.rewriter.constants.each do |ast,constant|
      output.puts "#{constant}\t#{ast}"
    end
    close(output)
  end
  
  # If no settable cells have been specified, then we assume that
  # all value cells should be settable if they are referenced by
  # any other forumla.
  def create_a_good_set_of_cells_that_should_be_settable_at_runtime
    references = all_formulae
    counter = CountFormulaReferences.new
    count = counter.count(references)

    count.each do |sheet,keys|
      keys.each do |ref,count|
        next unless count >= 1
        ast = references[sheet][ref]
        next unless ast
        if [:blank,:number,:null,:string,:shared_string,:constant,:percentage,:error,:boolean_true,:boolean_false].include?(ast.first)
          @cells_that_can_be_set_at_runtime[sheet] ||= []
          @cells_that_can_be_set_at_runtime[sheet] << ref.upcase
        end
      end
    end    
  end
  
  # UTILITY FUNCTIONS

  def settable(name)
    settable_refs = @cells_that_can_be_set_at_runtime[name]    
    if settable_refs
      lambda { |ref| (settable_refs == :all) ? true : settable_refs.include?(ref.upcase) } 
    else
      lambda { |ref| false }
    end
  end
  
  def gettable(name)
    if @cells_to_keep
      gettable_refs = @cells_to_keep[name]
      if gettable_refs
        lambda { |ref| (gettable_refs == :all) ? true : gettable_refs.include?(ref.upcase) }
      else
        lambda { |ref| false }
      end
    else
      lambda { |ref| true }
    end
  end
    
  def all_formulae
    references = {}
    worksheets do |name,xml_filename|
      r = references[name] = {}
      i = input([name,'Formulae'])
      i.lines do |line|
        line =~ /^(.*?)\t(.*)$/
        ref, ast = $1, $2
        r[$1] = eval($2)
      end
    end 
    references
  end
  
  def c_name_for_worksheet_name(name)
    unless @worksheet_names
      w = input('Worksheet C names')
      @worksheet_names = Hash[w.readlines.map { |line| line.split("\t").map { |a| a.strip }}]
      close(w)
    end
    @worksheet_names[name]
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
    
    @worksheet_filenames.each do |name, filename|
      block.call(name, filename)
    end
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
  
  def ruby_module_name
    @ruby_module_name = output_name.sub(/^[a-z\d]*/) { $&.capitalize }
    @ruby_module_name = @ruby_module_name.gsub(/(?:_|(\/))([a-z\d]*)/i) { "#{$1}#{$2.capitalize}" }.gsub('/', '::')
    @ruby_module_name
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
