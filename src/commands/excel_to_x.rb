# coding: utf-8
require 'fileutils'
require 'logger'
require_relative '../excel_to_code'

# FIXME: Correct case for all worksheet references
# FIXME: Correct case and $ stripping from all cell references
# FIXME: Replacing with c compatible names everywhere

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

  # Optional attribute. Specifies which named references to be turned into setters
  #
  # Should be an array of strings. Each string is a named reference. Case sensitive.
  # To specify a named reference scoped to a worksheet, use ['worksheet', 'named reference'] instead
  # of a string.
  #
  # Alternatively, can se to :where_possible to create setters for named references that point to setable cells
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
  # Should be an array of strings. Each string is a named reference. Case sensitive.
  #
  # To specify a named reference scoped to a worksheet, use ['worksheet', 'named reference'] instead
  # of a string.
  #
  # Alternatively, can specify :all to keep all named references
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
    if cells_that_can_be_set_at_runtime.is_a?(Hash)
      cells_that_can_be_set_at_runtime.keys.each do |sheet|
        next unless cells_that_can_be_set_at_runtime[sheet].is_a?(Array)
        cells_that_can_be_set_at_runtime[sheet] = cells_that_can_be_set_at_runtime[sheet].map { |reference| reference.gsub('$','').upcase }
      end
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
    
    # This turns named references that are specified as getters and setters
    # into a series of required cell references
    transfer_named_references_to_keep_into_cells_to_keep
    transfer_named_references_that_can_be_set_at_runtime_into_cells_that_can_be_set_at_runtime
    
    # These perform some translations to simplify the excel
    # Including:
    # * Turning row and column references (e.g., A:A) to areas, based on the size of the worksheet
    # * Turning range references (e.g., A1:B2) into array litterals (e.g., {A1,B1;A2,B2})
    # * Turning shared formulae into a series of conventional formulae
    # * Turning array formulae into a series of conventional formulae
    # * Mergining all the different types of formulae and values into a single file
    rewrite_worksheets
    
    # These perform a series of transformations to the information
    # with the intent of removing any redundant calculations
    # that are in the excel.
    simplify # Replacing shared strings and named references with their actual values, tidying arithmetic

    # In case this hasn't been set by the user
    if @cells_that_can_be_set_at_runtime.empty?
      log.info "Creating a good set of cells that should be settable"
      @cells_that_can_be_set_at_runtime = a_good_set_of_cells_that_should_be_settable_at_runtime
    end

    if named_references_that_can_be_set_at_runtime == :where_possible
      work_out_which_named_references_can_be_set_at_runtime
    end

    filter_named_references

    replace_formulae_with_their_results
    inline_formulae_that_are_only_used_once
    remove_any_cells_not_needed_for_outputs
    separate_formulae_elements
    replace_values_with_constants
    create_sorted_references_to_test

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
      parsed = Formula.parse(reference)
      if parsed
        @named_references[name] = parsed.to_ast[1]
      else
        $stderr.puts "Named reference #{name} #{reference} not parsed"
        exit
      end
    end
    # Replace A1:B2 with [A1, A2, B1, B2]
    rewriter = ReplaceRangesWithArrayLiteralsAst.new
    @named_references.each do |name, reference|
      @named_references[name] = rewriter.map(reference)
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
        @worksheet_xmls[name] = worksheet_xml
      end
    end
    # FIXME: Extract this and put it at the end ?
    @worksheet_c_names = {}
    worksheet_rids.keys.each do |excel_worksheet_name|
      @worksheet_c_names[excel_worksheet_name] = c_name_for(excel_worksheet_name)
    end
  end

  def c_name_for(name)
    @c_names_assigned ||= {}
    return @c_names_assigned.invert.fetch(name) if @c_names_assigned.has_value?(name)
    c_name = name.downcase.gsub(/[^a-z0-9]+/,'_') # Make it lowercase, replace anything that isn't a-z or 0-9 with underscores
    c_name = "s"+c_name if c_name[0] !~ /[a-z]/ # Can't start with a number. If it does, but an 's' in front (so 2010 -> s2010)
    c_name = c_name + "2" if @c_names_assigned.has_key?(c_name) # Add a number at the end if the c_name has already been used
    c_name.succ! while @c_names_assigned.has_key?(c_name)
    @c_names_assigned[c_name] = name
    c_name
  end

  # We want a central list of the maximum extent of each worksheet
  # so that we can convert column (e.g., C:F) and row (e.g., 13:18)
  # references into equivalent area references (e.g., C1:F30)
  def extract_dimensions_from_worksheets 
    log.info "Starting to extract dimensions from worksheets"  
    @worksheets_dimensions = {}
    extractor = ExtractWorksheetDimensions.new
    worksheets do |name, xml_filename|
      xml(xml_filename) do |i| 
        @worksheets_dimensions[name] = extractor.extract(i)
      end
      # FIXME: Should this actual return WorksheetDimension objects? rather than text ranges?
    end
  end
  
  # For each worksheet, this makes four passes through the xml
  # 1. Extract the values of each cell into @values
  # 2. Extract all the cells which are simple formulae into @formulae_simple
  # 3. Extract all the cells which use shared formulae
  # 4. Extract all the cells which are part of array formulae
  # 
  # It then looks at the relationship file and extracts any tables
  def extract_data_from_worksheets
    # All are hashes of the format ["SheetName", "A1"] => [:number, "1"]
    @values = {}
    @formulae_simple = {}
    @formulae_shared = {}
    @formulae_shared_targets = {}
    @formulae_array = {}
    # This one has a series of table references
    # FIXME: Should it actually have Table objects?
    @tables = {}
    
    # Loop through the worksheets
    # FIXME: make xml_filename be the IO object?
    worksheets do |name, xml_filename|
      log.info "Extracting data from #{name}"
      
      # ast
      xml(xml_filename) do |i|
        @values.merge! ExtractValues.extract(name, i)
      end
      
       # ast
      xml(xml_filename) do |i| 
        @formulae_simple.merge! ExtractSimpleFormulae.extract(name, i) 
      end
      
      # [shared_range, shared_identifier, ast]
      xml(xml_filename) do |i| 
        @formulae_shared.merge! ExtractSharedFormulae.extract(name, i) 
      end
      # shared_identifier
      xml(xml_filename) do |i| 
        @formulae_shared_targets.merge! ExtractSharedFormulaeTargets.extract(name, i)
      end
      #  [array_range, ast]
      xml(xml_filename) do |i| 
        @formulae_array.merge! ExtractArrayFormulae.extract(name, i)
      end
      
      extract_tables_for_worksheet(name,xml_filename)
    end
  end
  
  # To extract a table we need to look in the worksheet for table references
  # then we look in the relationships file for the filename that matches that
  # reference and contains the table data. Then we consolidate all the data
  # from individual table files into a single table file for the worksheet.
  def extract_tables_for_worksheet(name, xml_filename)
    table_rids = []
    xml(xml_filename) do |i| 
      table_rids = ExtractWorksheetTableRelationships.extract(i)
    end

    xml_for_rids = {}
    xml(File.join('worksheets','_rels',"#{File.basename(xml_filename)}.rels")) do |i|
      xml_for_rids = ExtractRelationships.extract(i)
    end

    table_rids.each do |rid| 
      # FIXME: Extract actual Table objects?
      xml(File.join('worksheets', xml_for_rids[rid])) do |i|
        @tables.merge! ExtractTable.extract(name, i)
      end
    end
  end
  
  def rewrite_worksheets
    rewrite_row_and_column_references
    rewrite_shared_formulae
    rewrite_array_formulae
    rewrite_values
    combine_formulae_files
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
  
  def rewrite_shared_formulae
    log.info "Rewriting shared formulae"
    @formulae_shared = RewriteSharedFormulae.rewrite( @formulae_shared, @formulae_shared_targets)
    # FIXME: Could now nil off the @formula_shared_targets ?
  end
  
  def rewrite_array_formulae
    log.info "Rewriting array formulae"
    # FIMXE: Refactor this

    # Replace the named references in the array formulae
    named_references = NamedReferences.new(@named_references)
    named_reference_replacer = ReplaceNamedReferencesAst.new(named_references) 

    @formulae_array.each do |ref, details|
      named_reference_replacer.default_sheet_name = ref.first
      named_reference_replacer.map(details.last)
    end
    
    # FIXME: Refactor
    table_objects = {}
    @tables.each do |name, details|
      table_objects[name.downcase] = Table.new(name, *details)
    end
    
    table_reference_replacer = ReplaceTableReferenceAst.new(table_objects)

    @formulae_array.each do |ref, details|
      table_reference_replacer.worksheet = ref.first
      table_reference_replacer.referring_cell = ref.last
      table_reference_replacer.map(details.last)
    end
      
    simplify_arithmetic_replacer = SimplifyArithmeticAst.new
    @formulae_array.each do |ref, details|
      simplify_arithmetic_replacer.map(details.last)
    end
    
    replace_ranges_with_array_literals_replacer = ReplaceRangesWithArrayLiteralsAst.new
    @formulae_array.each do |ref, details|
      replace_ranges_with_array_literals_replacer.map(details.last)
    end

    expand_array_formulae_replacer = AstExpandArrayFormulae.new
    @formulae_array.each do |ref, details|
      expand_array_formulae_replacer.map(details.last)
    end

    @formulae_array = RewriteArrayFormulae.rewrite(@formulae_array)
  end

  def rewrite_values
    log.info "Rewriting values"
    r = ReplaceSharedStringAst.new(@shared_strings)
    @values.each do |ref, ast|
      r.map(ast)
    end
  end
  
  def combine_formulae_files
    log.info "Combining formula files"

    @formulae = required_references
    # We dup this to avoid the values being replaced when manipulating formulae
    @values.each do |ref, value|
      @formulae[ref] = value.dup
    end
    @formulae.merge! @formulae_shared
    @formulae.merge! @formulae_array
    @formulae.merge! @formulae_simple
  end
  
  # This ensures that all gettable and settable values appear in the output
  # even if they are blank in the underlying excel
  def required_references
    log.info "Checking required references"
    required_refs = {}
    if @cells_that_can_be_set_at_runtime && @cells_that_can_be_set_at_runtime != :named_references_only
      @cells_that_can_be_set_at_runtime.each do |worksheet, refs|
        next if refs == :all
        refs.each do |ref|
          required_refs[[worksheet, ref]] = [:blank]
        end
      end
    end
    if @cells_to_keep
      @cells_to_keep.each do |worksheet, refs|
        next if refs == :all
        refs.each do |ref|
          required_refs[[worksheet, ref]] = [:blank]
        end
      end
    end
    required_refs
  end

  # This makes sure that cells_to_keep includes named_references_to_keep
  def transfer_named_references_to_keep_into_cells_to_keep
    log.info "Transfering named references to keep into cells to keep"
    return unless @named_references_to_keep
    @named_references_to_keep = @named_references.keys if @named_references_to_keep == :all
    @cells_to_keep ||= {}
    @named_references_to_keep.each do |name|
      ref = @named_references[name]
      if ref
        add_ref_to_hash(ref, @cells_to_keep)
      else
        log.warn "Named reference #{name} not found"
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
      ref = @named_references[name]
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
      cell = ref[2][1].gsub('$','')
      hash[sheet] ||= []
      return if hash[sheet] == :all
      hash[sheet] << cell unless hash[sheet].include?(cell)
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

  # This just checks which named references refer to cells that we have already declared as settable
  def work_out_which_named_references_can_be_set_at_runtime
    log.info "Working out which named references can be set at runtime"
    return unless @named_references_that_can_be_set_at_runtime
    return unless @named_references_that_can_be_set_at_runtime == :where_possible
    cells_that_can_be_set = @cells_that_can_be_set_at_runtime
    cells_that_can_be_set = a_good_set_of_cells_that_should_be_settable_at_runtime if cells_that_can_be_set == :named_references_only
    cells_that_can_be_set_due_to_named_reference = Hash.new { |h,k| h[k] = Array.new  }
    @named_references_that_can_be_set_at_runtime = []
    all_named_references = @named_references
    # FIXME can this be refactored with #add_ref_to_hash
    @named_references_to_keep.each do |name|
      ref = all_named_references[name]
      if ref.first == :sheet_reference
        sheet = ref[1]
        cell = ref[2][1].gsub('$','')
        s = cells_that_can_be_set[sheet]
        if s && s.include?(cell)
          @named_references_that_can_be_set_at_runtime << name 
          cells_that_can_be_set_due_to_named_reference[sheet] << cell
          cells_that_can_be_set_due_to_named_reference[sheet].uniq!
        end
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
            cells_that_can_be_set_due_to_named_reference[sheet] << cell
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
    log.info "Filtering named refernces to keep"
    @named_references_to_keep ||= []
    @named_references_that_can_be_set_at_runtime ||= []

    @named_references.each do |name, ref|
      if named_references_to_keep.include?(name) || named_references_that_can_be_set_at_runtime.include?(name)
        # FIXME: Refactor the c_name_for to closer to the writing?
        @named_references_to_keep << name
      end
    end

    @named_references.each do |name, ref|
      if named_references_that_can_be_set_at_runtime.include?(name)
        @named_references_that_can_be_set_at_runtime << name
      end
    end
  end
    
  def simplify(cells = @formulae)
    log.info "Simplifying cells"
    r = ReplaceSharedStringAst.new(@shared_strings)
    cells.each do |ref, ast|
      r.map(ast)
    end

    simplify_arithmetic_replacer = SimplifyArithmeticAst.new
    cells.each do |ref, ast|
      simplify_arithmetic_replacer.map(ast)
    end
      
    # Replace the named references in the array formulae
    named_references = NamedReferences.new(@named_references)
    named_reference_replacer = ReplaceNamedReferencesAst.new(named_references) 

    cells.each do |ref, ast|
      named_reference_replacer.default_sheet_name = ref.first
      named_reference_replacer.map(ast)
    end
    
    # FIXME: Refactor
    table_objects = {}
    @tables.each do |name, details|
      table_objects[name.downcase] = Table.new(name, *details)
    end
    
    table_reference_replacer = ReplaceTableReferenceAst.new(table_objects)

    cells.each do |ref, ast|
      table_reference_replacer.worksheet = ref.first
      table_reference_replacer.referring_cell = ref.last
      table_reference_replacer.map(ast)
    end
      
    replace_ranges_with_array_literals_replacer = ReplaceRangesWithArrayLiteralsAst.new
    cells.each do |ref, ast|
      replace_ranges_with_array_literals_replacer.map(ast)
    end

    replace_arithmetic_on_ranges_replacer = ReplaceArithmeticOnRangesAst.new
    cells.each do |ref, ast|
      replace_arithmetic_on_ranges_replacer.map(ast)
    end

    replace_arrays_with_single_cells_replacer = ReplaceArraysWithSingleCellsAst.new
    cells.each do |ref, ast|
      replace_arrays_with_single_cells_replacer.map(ast)
    end

    replace_string_joins_on_ranges_replacer = ReplaceStringJoinOnRangesAST.new
    cells.each do |ref, ast|
      replace_string_joins_on_ranges_replacer.map(ast)
    end

    r = RewriteCellReferencesToIncludeSheetAst.new
    cells.each do |ref, ast|
      r.worksheet = ref.first
      r.map(ast)
    end

    wrap_formulae_that_return_arrays_replacer = WrapFormulaeThatReturnArraysAndAReNotInArraysAst.new
    cells.each do |ref, ast|
      wrap_formulae_that_return_arrays_replacer.map(ast)
    end
  end
    
  def replace_formulae_with_their_results
    number_of_passes = 0
    begin 
      number_of_passes += 1
      @replacements_made_in_the_last_pass = 0
      replace_references_to_values_with_values
      replace_formulae_with_calculated_values
      replace_indirects_and_offsets
      log.info "Pass #{number_of_passes}: Made #{@replacements_made_in_the_last_pass} replacements"
      if number_of_passes > 20
        log.warn "Made more than 20 passes, so aborting"
        break
      end
    end while @replacements_made_in_the_last_pass > 0
  end
  
  # There is no support for INDIRECT or OFFSET in the ruby or c runtime
  # However, in many cases it isn't needed, because we can work
  # out the value of the indirect or OFFSET at compile time and eliminate it
  # First of all we replace any indirects where their values can be calculated at compile time with those
  # calculated values (e.g., INDIRECT("A"&1) can be turned into A1 and OFFSET(A1,1,1,2,2) can be turned into B2:C3)
  def replace_indirects_and_offsets
    log.info "Replacing indirects and offsets"

    references_that_need_updating = {}

    indirect_replacement = ReplaceIndirectsWithReferencesAst.new
    column_replacement = ReplaceColumnWithColumnNumberAST.new
    offset_replacement = ReplaceOffsetsWithReferencesAst.new

    @formulae.each do |ref, ast|
      if column_replacement.replace(ast)
        references_that_need_updating[ref] = ast
      end
      if offset_replacement.replace(ast)
        references_that_need_updating[ref] = ast
      end
      if indirect_replacement.replace(ast)
        references_that_need_updating[ref] = ast
      end
    end
    @replacements_made_in_the_last_pass += column_replacement.count_replaced
    @replacements_made_in_the_last_pass += offset_replacement.count_replaced
    @replacements_made_in_the_last_pass += indirect_replacement.count_replaced

    simplify(references_that_need_updating)
  end
  
  # If a formula's value can be calculated at compile time, it is replaced with its calculated value (e.g., 1+1 gets replaced with 2)
  def replace_formulae_with_calculated_values    
    log.info "Replacing formulae with calculated values"

    value_replacer = MapFormulaeToValues.new
    value_replacer.original_excel_filename = excel_file
    @formulae.each do |ref, ast|
      value_replacer.map(ast)
    end
    @replacements_made_in_the_last_pass += value_replacer.replacements_made_in_the_last_pass
  end

  # If a formula references a cell containing a value, the reference is replaced with the value (e.g., if A1 := 2 and A2 := A1 + 1 then becomes: A2 := 2 + 1)
  def replace_references_to_values_with_values
    log.info "Replacing references to values with values"
    
    inline_ast_decision = lambda do |sheet, cell, references|
      references_to_keep = @cells_that_can_be_set_at_runtime[sheet]
      if references_to_keep && (references_to_keep == :all || references_to_keep.include?(cell))
        false
      else
        ast = references[[sheet,cell]]
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
    
    r = InlineFormulaeAst.new
    r.references = @formulae
    r.inline_ast = inline_ast_decision
    
    @formulae.each do |ref, ast|
      # FIXME: Shouldn't need to wrap ref.fist in an array
      r.current_sheet_name = [ref.first]
      r.map(ast)
    end
    
    @replacements_made_in_the_last_pass += r.count_replaced
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
    p identifier.dependencies["Model"]["C154"]
    
    # Now we actually go ahead and remove the cells
    r = RemoveCells.new
    r.cells_to_keep = identifier.dependencies
    r.rewrite(@formulae)
    # Must remove the values as well, to avoid any tests being generated for cells that don't exist
    r.rewrite(@values)
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
    @formulae.each do |ref, ast|
      r.current_sheet_name = [ref.first]
      r.map(ast)
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
    repeated_elements = identifier.count(@formulae)
    
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

    # FIXME: This means that some common elements won't ever be called, becuase they are replaced by a longer common element. Should the common elements be merged first?
    ReplaceCommonElementsInFormulae.replace(@formulae, repeated_element_ast)

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

    @constants = r.constants.invert
  end
  
  # If nothing has been specified in named_references_that_can_be_set_at_runtime 
  # or in cells_that_can_be_set_at_runtime, then we assume that
  # all value cells should be settable if they are referenced by
  # any other forumla.
  def a_good_set_of_cells_that_should_be_settable_at_runtime
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
    @worksheet_c_names[name]
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

end
