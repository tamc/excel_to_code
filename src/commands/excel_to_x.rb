# coding: utf-8
require 'fileutils'
require_relative '../excel_to_code'

# Used to throw normally fatal errors
class ExcelToCodeException < Exception; end

class ExcelToX
  
  # Required attribute. The source excel file. This must be .xlsx not .xls
  attr_accessor :excel_file
  
  # Optional attribute. The output directory.
  #  If not specified, will be '#{excel_file_name}/c'
  attr_accessor :output_directory
  
  # Optional attribute. The name of the resulting c file (and associated ruby ffi module). Defaults to excelspreadsheet
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
      cells_that_can_be_set_at_runtime[sheet] = cells_that_can_be_set_at_runtime[sheet].map { |reference| reference.gsub('$','').downcase }
    end

    # Make sure that all the cell names are downcase and don't have any $ in them
    if cells_to_keep
      cells_to_keep.keys.each do |sheet|
        next unless cells_to_keep[sheet].is_a?(Array)
        cells_to_keep[sheet] = cells_to_keep[sheet].map { |reference| reference.gsub('$','').downcase }
      end
    end  
    
    # Make sure the relevant directories exist
    self.excel_file = File.expand_path(excel_file)
    self.output_directory = File.expand_path(output_directory)
  end
  
  def go!
    # This sorts out the attributes
    set_defaults
    
    # These turn the excel into a more accesible format
    sort_out_output_directories
    unzip_excel
    
    # These get all the information out of the excel and put
    # into a useful format
    extract_data_from_workbook
    extract_data_from_worksheets
    merge_table_files
    
    rewrite_worksheets
    
    # These perform a series of transformations to the information
    # with the intent of removing any redundant calculations
    # that are in the excel
    simplify_worksheets
    optimise_and_replace_indirect_loop
    replace_blanks
    remove_any_cells_not_needed_for_outputs
    inline_formulae_that_are_only_used_once
    separate_formulae_elements
    replace_values_with_constants
    
    # This actually creates the code (implemented in subclasses)
    write_code
    
    # These compile and run the code version of the excel
    compile_code
    run_tests
    
    puts
    puts "The generated code is available in #{File.join(output_directory)}"
  end
  
  def sort_out_output_directories    
    FileUtils.mkdir_p(output_directory)
    FileUtils.mkdir_p(xml_directory)
    FileUtils.mkdir_p(intermediate_directory)
  end
  
  def unzip_excel
    puts `rm -fr '#{xml_directory}'`
    puts `unzip -uo '#{excel_file}' -d '#{xml_directory}'`
  end

  def extract_data_from_workbook
    extract_shared_strings
    extract_named_references
    extract_worksheet_names
    extract_dimensions_from_worksheets
  end
  
  def extract_shared_strings
    if File.exists?(File.join(xml_directory,'xl','sharedStrings.xml'))
      extract ExtractSharedStrings, 'sharedStrings.xml', 'shared_strings'
    else
      FileUtils.touch(File.join(intermediate_directory,'shared_strings'))
    end
  end
  
  def extract_named_references
    extract ExtractNamedReferences, 'workbook.xml', 'named_references'
    rewrite RewriteFormulaeToAst, 'named_references', 'named_references.ast'
  end
  
  def extract_worksheet_names
    extract ExtractWorksheetNames, 'workbook.xml', 'worksheet_names_without_filenames'
    extract ExtractRelationships, File.join('_rels','workbook.xml.rels'), 'workbook_relationships'
    rewrite RewriteWorksheetNames, 'worksheet_names_without_filenames', 'workbook_relationships', 'worksheet_names'
    rewrite MapSheetNamesToCNames, 'worksheet_names', 'worksheet_c_names'
  end
  
  def extract_dimensions_from_worksheets    
    dimension_file = intermediate('dimensions')
    worksheets("Extracting dimensions") do |name,xml_filename|
      dimension_file.write name
      dimension_file.write "\t"
      extract ExtractWorksheetDimensions, File.open(xml_filename,'r'), dimension_file 
    end
    dimension_file.close
  end
  
  def extract_data_from_worksheets
    worksheets("Initial data extract") do |name,xml_filename|
      worksheet_directory = File.join(intermediate_directory,name)
      worksheet_xml = File.open(xml_filename,'r')
      
      worksheet_xml.rewind
      extract ExtractValues, worksheet_xml, File.join(name,'values')
      rewrite RewriteValuesToAst, File.join(name,'values'), File.join(name,'values.ast')
      
      worksheet_xml.rewind
      extract ExtractSimpleFormulae, worksheet_xml, File.join(name,'simple_formulae')
      rewrite RewriteFormulaeToAst,  File.join(name,'simple_formulae'),  File.join(name,'simple_formulae.ast')

      worksheet_xml.rewind
      extract ExtractSharedFormulae, worksheet_xml, File.join(name,'shared_formulae')
      rewrite RewriteFormulaeToAst,  File.join(name,'shared_formulae'),  File.join(name,'shared_formulae.ast')

      worksheet_xml.rewind
      extract ExtractArrayFormulae, worksheet_xml, File.join(name,'array_formulae')
      rewrite RewriteFormulaeToAst,  File.join(name,'array_formulae'),  File.join(name,'array_formulae.ast')
      
      worksheet_xml.rewind
      extract ExtractWorksheetTableRelationships, worksheet_xml, File.join(name,'table_rids')
      if File.exists?(File.join(xml_directory,'xl','worksheets','_rels',"#{File.basename(xml_filename)}.rels"))
        extract_tables(name,xml_filename)
      else
        fake_extract_tables(name,xml_filename)
      end
      close(worksheet_xml)
    end
  end
  
  def extract_tables(name,xml_filename)
    extract ExtractRelationships, File.join('worksheets','_rels',"#{File.basename(xml_filename)}.rels"), File.join(name,'relationships')
    rewrite RewriteRelationshipIdToFilename, File.join(name,'table_rids'), File.join(name,'relationships'), File.join(name,'table_filenames')
    tables = intermediate(name,'tables')
    table_extractor = ExtractTable.new(name)
    table_filenames = input(name,'table_filenames')
    table_filenames.lines.each do |line|
      extract table_extractor, File.join('worksheets',line.strip), tables
    end
    close(tables,table_filenames)
  end
    
  def fake_extract_tables(name,xml_filename)
      a = intermediate(name,'relationships')
      b = intermediate(name,'table_filenames')
      c = intermediate(name,'tables')
      close(a,b,c)
  end
  
  def rewrite_worksheets
    worksheets("Initial rewrite of references and formulae") do |name,xml_filename|
        rewrite_row_and_column_references(name,xml_filename)
        rewrite_shared_formulae(name,xml_filename)
        rewrite_array_formulae(name,xml_filename)
        combine_formulae_files(name,xml_filename)
    end
  end
  
  def rewrite_row_and_column_references(name,xml_filename)
    dimensions = input('dimensions')
    %w{simple_formulae.ast shared_formulae.ast array_formulae.ast}.each do |file|
      dimensions.rewind
      i = input(name,file)
      o = intermediate(name,"#{file}-nocols")
      RewriteWholeRowColumnReferencesToAreas.rewrite(i,name, dimensions, o)
      close(i,o)
    end
    dimensions.close
  end
  
  def rewrite_shared_formulae(name,xml_filename)
    i = input(name,'shared_formulae.ast-nocols')
    o = intermediate(name,"shared_formulae-expanded.ast")
    RewriteSharedFormulae.rewrite(i,o)
    close(i,o)
  end
  
  def rewrite_array_formulae(name,xml_filename)
    r = ReplaceNamedReferences.new
    r.sheet_name = name
    replace r, File.join(name,'array_formulae.ast-nocols'), 'named_references.ast', File.join(name,"array_formulae1.ast")

    r = ReplaceTableReferences.new
    r.sheet_name = name    
    replace r, File.join(name,'array_formulae1.ast'), 'all_tables', File.join(name,"array_formulae2.ast")
    replace SimplifyArithmetic, File.join(name,'array_formulae2.ast'), File.join(name,'array_formulae3.ast')
    replace ReplaceRangesWithArrayLiterals, File.join(name,"array_formulae3.ast"), File.join(name,"array_formulae4.ast")
    rewrite RewriteArrayFormulaeToArrays, File.join(name,"array_formulae4.ast"), File.join(name,"array_formulae5.ast")
    rewrite RewriteArrayFormulae, File.join(name,'array_formulae5.ast'), File.join(name,"array_formulae-expanded.ast")
  end
  
  def combine_formulae_files(name,xml_filename)
    values = File.join(name,'values.ast')
    shared_formulae = File.join(name,"shared_formulae-expanded.ast")
    array_formulae = File.join(name,"array_formulae-expanded.ast")
    simple_formulae = File.join(name,"simple_formulae.ast-nocols")
    output = File.join(name,'formulae.ast')
    rewrite RewriteMergeFormulaeAndValues, values, shared_formulae, array_formulae, simple_formulae, output
  end
  
  def merge_table_files
    tables = []
    worksheets("Merging table files") do |name,xml_filename|
      tables << File.join(name,'tables')
    end
    if run_in_memory
      o = intermediate("all_tables")
      tables.each do |t|
        i = input(t)
        o.print i.string
        close(i)
      end
      close(o)
    else
      `sort #{tables.map { |t| " '#{File.join(intermediate_directory,t)}' "}.join} > #{File.join(intermediate_directory,'all_tables')}`
    end
  end
  
  def simplify_worksheets
    worksheets("Simplifying") do |name,xml_filename|
      replace SimplifyArithmetic, File.join(name,'formulae.ast'), File.join(name,'formulae_simple_arithmetic.ast')
      
      replace ReplaceSharedStrings, File.join(name,'formulae_simple_arithmetic.ast'), 'shared_strings', File.join(name,"formulae_no_shared_strings.ast")
      replace ReplaceSharedStrings, File.join(name,'values.ast'), 'shared_strings', File.join(name,"values_no_shared_strings.ast")
      
      r = ReplaceNamedReferences.new
      r.sheet_name = name
      replace r, File.join(name,'formulae_no_shared_strings.ast'), 'named_references.ast', File.join(name,"formulae_no_named_references.ast")

      r = ReplaceTableReferences.new
      r.sheet_name = name
      replace r, File.join(name,'formulae_no_named_references.ast'), 'all_tables', File.join(name,"formulae_no_table_references.ast")
      
      replace ReplaceRangesWithArrayLiterals, File.join(name,"formulae_no_table_references.ast"), File.join(name,"formulae_no_ranges.ast")
    end
  end
  
  def replace_blanks
    references = all_formulae("formulae_no_indirects_optimised.ast")
    r = ReplaceBlanks.new
    r.references = references
    worksheets("Replacing blanks") do |name,xml_filename|
      r.default_sheet_name = name
      replace r, File.join(name,"formulae_no_indirects_optimised.ast"),File.join(name,"formulae_no_blanks.ast")
    end
  end
  
  def optimise_and_replace_indirect_loop
    number_of_loops = 4
    1.upto(number_of_loops) do |pass|
      puts "Optimise and replace indirects pass #{pass}"
      start = pass == 1 ? "formulae_no_ranges.ast" : "optimse-output-#{pass-1}.ast"
      finish = pass == number_of_loops ? "formulae_no_indirects_optimised.ast" : "optimse-output-#{pass}.ast"
      replace_indirects(start,"replace-indirect-output-#{pass}.ast","replace-indirect-working-#{pass}-")
      optimise_sheets("replace-indirect-output-#{pass}.ast",finish,"optimse-working-#{pass}-")
    end
  end
  
  def replace_indirects(start_filename,finish_filename,basename)
    worksheets("Replacing indirects") do |name,xml_filename|
      counter = 1
      replace ReplaceIndirectsWithReferences, File.join(name,start_filename), File.join(name,"#{basename}#{counter+1}.ast")
      counter += 1

      r = ReplaceNamedReferences.new
      r.sheet_name = name
      replace r, File.join(name,"#{basename}#{counter}.ast"), 'named_references.ast', File.join(name,"#{basename}#{counter+1}.ast")
      counter += 1
      
      r = ReplaceTableReferences.new
      r.sheet_name = name
      replace r, File.join(name,"#{basename}#{counter}.ast"), 'all_tables', File.join(name,"#{basename}#{counter+1}.ast")
      counter += 1
      
      replace ReplaceRangesWithArrayLiterals, File.join(name,"#{basename}#{counter}.ast"), File.join(name,"#{basename}#{counter+1}.ast")
      counter += 1
      replace ReplaceArraysWithSingleCells, File.join(name,"#{basename}#{counter}.ast"), File.join(name,"#{basename}#{counter+1}.ast")
      counter += 1
      
      # Finally, create the output file
      i = File.join(intermediate_directory,name,"#{basename}#{counter}.ast")
      o = File.join(intermediate_directory,name,finish_filename)
      if run_in_memory
        @files[o] = @files[i]
      else
        `cp '#{i}' '#{o}'`
      end
    end
  end
  
  def optimise_sheets(start_filename,finish_filename,basename)
    counter = 1
    
    # Setup start
    worksheets("Setting up for optimise -#{counter}") do |name|
      i = File.join(intermediate_directory,name,start_filename)
      o = File.join(intermediate_directory,name,"#{basename}#{counter}.ast")
      if run_in_memory
        @files[o] = @files[i]
      else
        `cp '#{i}' '#{o}'`
      end
    end
    
    worksheets("Replacing with calculated values #{counter}-#{counter+1}") do |name,xml_filename|
      replace ReplaceFormulaeWithCalculatedValues, File.join(name,"#{basename}#{counter}.ast"), File.join(name,"#{basename}#{counter+1}.ast")
    end
    counter += 1
    Process.waitall
      
    references = all_formulae("#{basename}#{counter}.ast")
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
    
    worksheets("Inlining formulae #{counter}-#{counter+1}") do |name,xml_filename|
      r.default_sheet_name = name
      replace r, File.join(name,"#{basename}#{counter}.ast"), File.join(name,"#{basename}#{counter+1}.ast")
    end
    counter += 1
    Process.waitall
    
    # Finish
    worksheets("Moving sheets #{counter}-") do |name|
      o = File.join(intermediate_directory,name,finish_filename)
      i = File.join(intermediate_directory,name,"#{basename}#{counter}.ast")
      if run_in_memory
        @files[o] = @files[i]
      else
        `cp '#{i}' '#{o}'`
      end
    end
  end
  
  def remove_any_cells_not_needed_for_outputs(formula_in = "formulae_no_blanks.ast", formula_out = "formulae_pruned.ast", values_in = "values_no_shared_strings.ast", values_out = "values_pruned.ast")
    if cells_to_keep && !cells_to_keep.empty?
      identifier = IdentifyDependencies.new
      identifier.references = all_formulae(formula_in)
      cells_to_keep.each do |sheet_to_keep,cells_to_keep|
        if cells_to_keep == :all
          identifier.add_depedencies_for(sheet_to_keep)
        elsif cells_to_keep.is_a?(Array)
          cells_to_keep.each do |cell|
            identifier.add_depedencies_for(sheet_to_keep,cell)
          end
        end
      end
      r = RemoveCells.new
      worksheets("Removing cells") do |name,xml_filename|
          r.cells_to_keep = identifier.dependencies[name]
          rewrite r, File.join(name, formula_in), File.join(name, formula_out)
          rewrite r, File.join(name, values_in), File.join(name, values_out)
      end
    else
      worksheets do |name,xml_filename|
        i = File.join(intermediate_directory,name, formula_in)
        o = File.join(intermediate_directory,name, formula_out)
        if run_in_memory
          @files[o] = @files[i]
        else
          `cp '#{i}' '#{o}'`
        end
        i = File.join(intermediate_directory,name, values_in)
        o = File.join(intermediate_directory,name, values_out)
        if run_in_memory
          @files[o] = @files[i]
        else
          `cp '#{i}' '#{o}'`
        end
      end
    end
  end
  
  def inline_formulae_that_are_only_used_once
    references = all_formulae("formulae_pruned.ast")
    counter = CountFormulaReferences.new
    count = counter.count(references)
    
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
    
    worksheets("Inlining formulae") do |name,xml_filename|
      r.default_sheet_name = name
      replace r, File.join(name,"formulae_pruned.ast"), File.join(name,"formulae_inlined.ast")
    end
    
    remove_any_cells_not_needed_for_outputs("formulae_inlined.ast", "formulae_inlined_pruned.ast", "values_pruned.ast", "values_pruned2.ast")
  end
  
  def separate_formulae_elements
    # First we add the sheet to all references, so that we can then look for common elements accross worksheets
    r = RewriteCellReferencesToIncludeSheet.new
    worksheets("Adding the sheet to all references") do |name,xml_filename|
      r.worksheet = name
      rewrite r, File.join(name,"formulae_inlined_pruned.ast"), File.join(name,"formulae_inlined_pruned_with_sheets.ast") 
    end
    
    references = all_formulae("formulae_inlined_pruned_with_sheets.ast")
    identifier = IdentifyRepeatedFormulaElements.new
    repeated_elements = identifier.count(references)
    repeated_elements.delete_if do |element,count|
      count < 2
    end
    o = intermediate('common-elements-1.ast')
    i = 0
    repeated_elements.each do |element,count|
      o.puts "common#{i}\t#{element}"
      i = i + 1
    end
    close(o)
    
    worksheets("Replacing repeated elements") do |name,xml_filename|
      replace ReplaceCommonElementsInFormulae, File.join(name,"formulae_inlined_pruned_with_sheets.ast"), "common-elements-1.ast", File.join(name,"formulae_inlined_pruned_replaced-1.ast")
    end
  end
  
  def replace_values_with_constants
    r = ReplaceValuesWithConstants.new  
    worksheets("Replacing values with constants") do |name,xml_filename|
      i = input(name,"formulae_inlined_pruned_replaced-1.ast")
      o = intermediate(name,"formulae_inlined_pruned_replaced.ast")
      r.replace(i,o)
      close(i,o)
    end
    
    puts "Replacing values with constants in common elements"
    i = input("common-elements-1.ast")
    o = intermediate("common-elements.ast")
    r.replace(i,o)
    close(i,o)
    
    puts "Writing out constants"
    co = intermediate("value_constants.ast")
    r.rewriter.constants.each do |ast,constant|
      co.puts "#{constant}\t#{ast}"
    end
    close(co)
  end
  
  # UTILITY FUNCTIONS

  def settable(name)
    settable_refs = @cells_that_can_be_set_at_runtime[name]    
    if settable_refs
      lambda { |ref| (settable_refs == :all) ? true : settable_refs.include?(ref) } 
    else
      lambda { |ref| false }
    end
  end
  
  def gettable(name)
    if @cells_to_keep
      gettable_refs = @cells_to_keep[name]
      if gettable_refs
        lambda { |ref| (gettable_refs == :all) ? true : gettable_refs.include?(ref) }
      else
        lambda { |ref| false }
      end
    else
      lambda { |ref| true }
    end
  end
    
  def all_formulae(filename)
    references = {}
    worksheets do |name,xml_filename|
      r = references[name] = {}
      i = input(name,filename)
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
      w = input("worksheet_c_names")
      @worksheet_names = Hash[w.readlines.map { |line| line.split("\t").map { |a| a.strip }}]
      close(w)
    end
    @worksheet_names[name]
  end
    
  def worksheets(message = "Processing",&block)
    input('worksheet_names').lines.each do |line|
      name, filename = *line.split("\t")
      filename = File.expand_path(File.join(xml_directory,'xl',filename.strip))
      puts "#{message} #{name}"
      block.call(name, filename)
    end
  end
    
  def extract(_klass,xml_name,output_name)
    i = xml_name.is_a?(String) ? xml(xml_name) : xml_name
    o = output_name.is_a?(String) ? intermediate(output_name) : output_name
    _klass.extract(i,o)
    if xml_name.is_a?(String)
      close(i)
    end
    if output_name.is_a?(String)
      close(o)
    end
  end
  
  def rewrite(_klass,*args)
    o = intermediate(args.pop)
    inputs = args.map { |name| input(name) }
    _klass.rewrite(*inputs,o)
    close(*inputs,o)
  end
  
  def replace(_klass,*args)
    o = intermediate(args.pop)
    inputs = args.map { |name| input(name) }
    _klass.replace(*inputs,o)
    close(*inputs,o)
  end
  
  def xml(*args)
    File.open(File.join(xml_directory,'xl',*args),'r')
  end
  
  def input(*args)
    filename = File.join(intermediate_directory,*args)
    if run_in_memory
      io = StringIO.new(@files[filename].string,'r')
    else
      File.open(filename,'r')
    end
  end
  
  def intermediate(*args)
    filename = File.join(intermediate_directory,*args)
    if run_in_memory
      @files ||= {}
      @files[filename] = StringIO.new("",'w')
    else
      FileUtils.mkdir_p(File.dirname(filename))
      File.open(filename,'w')
    end
  end
  
  def output(*args)
    File.open(File.join(output_directory,*args),'w')
  end
  
  def close(*args)
    args.map do |f|
      next if f.is_a?(StringIO)
      next if f.is_a?(String)
      f.close
    end
  end
  
end