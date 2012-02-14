# coding: utf-8

require 'fileutils'
require_relative '../util'
require_relative '../excel'
require_relative '../extract'
require_relative '../rewrite'
require_relative '../simplify'
require_relative '../compile'

class ExcelToRuby
  
  attr_accessor :excel_file, :output_directory, :xml_dir, :compiled_module_name, :values_that_can_be_set_at_runtime, :outputs_to_keep
  
  def initialize
    @values_that_can_be_set_at_runtime ||= {}
  end
  
  def go!
    self.excel_file = File.expand_path(excel_file)
    self.output_directory = File.expand_path(output_directory)
    self.xml_dir = File.join(output_directory,'xml')
  
    sort_out_output_directories
    unzip_excel
    process_workbook
    extract_worksheets
    Process.waitall
    merge_table_files
    rewrite_worksheets
    Process.waitall
    simplify_worksheets
    Process.waitall
    optimise_and_replace_indirect_loop
    Process.waitall
    replace_blanks
    Process.waitall
    remove_any_cells_not_needed_for_outputs
    Process.waitall
    inline_formulae_that_are_only_used_once
    Process.waitall
    separate_formulae_elements
    Process.waitall
    compile_workbook
    compile_worksheets
    Process.waitall
  end
  
  def sort_out_output_directories    
    FileUtils.mkdir_p(File.join(output_directory,'intermediate'))
    FileUtils.mkdir_p(File.join(output_directory,'ruby','worksheets'))
    FileUtils.mkdir_p(File.join(output_directory,'ruby','tests'))
  end
  
  def unzip_excel
    puts `unzip -uo '#{excel_file}' -d '#{xml_dir}'`
  end

  def process_workbook    
    extract ExtractSharedStrings, 'sharedStrings.xml', 'shared_strings'
    
    extract ExtractNamedReferences, 'workbook.xml', 'named_references'
    rewrite RewriteFormulaeToAst, 'named_references', 'named_references.ast'
    
    extract ExtractWorksheetNames, 'workbook.xml', 'worksheet_names_without_filenames'
    extract ExtractRelationships, File.join('_rels','workbook.xml.rels'), 'workbook_relationships'
    rewrite RewriteWorksheetNames, 'worksheet_names_without_filenames', 'workbook_relationships', 'worksheet_names'
    rewrite MapSheetNamesToRubyNames, 'worksheet_names', 'worksheet_ruby_names'
    
    extract_dimensions_from_worksheets
  end
  
  # Extracts each worksheets values and formulas
  def extract_worksheets
    worksheets("Initial data extract") do |name,xml_filename|
      #fork do
        $0 = "ruby initial extract #{name}"
        initial_extract_from_worksheet(name,xml_filename)
      #end
    end
  end

  # Extracts the dimensions of each worksheet and puts them in a single file  
  def extract_dimensions_from_worksheets    
    dimension_file = output('dimensions')
    worksheets("Extracting dimensions") do |name,xml_filename|
      dimension_file.write name
      dimension_file.write "\t"
      extract ExtractWorksheetDimensions, File.open(xml_filename,'r'), dimension_file 
    end
    dimension_file.close
  end
  
  def rewrite_worksheets
    worksheets("Initial rewrite of references and formulae") do |name,xml_filename|
      #fork do 
        rewrite_row_and_column_references(name,xml_filename)
        rewrite_shared_formulae(name,xml_filename)
        rewrite_array_formulae(name,xml_filename)
        combine_formulae_files(name,xml_filename)
      #end
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
      table_filenames = input(name,'table_filenames')
      table_filenames.lines.each do |line|
        extract table_extractor, File.join('worksheets',line.strip), tables
      end
      close(tables,table_filenames)
    else
      FileUtils.touch File.join(output_directory,'intermediate',name,'relationships')
      FileUtils.touch File.join(output_directory,'intermediate',name,'table_filenames')      
      FileUtils.touch File.join(output_directory,'intermediate',name,'tables')      
    end
    close(worksheet_xml)
  end
  
  def merge_table_files
    tables = []
    worksheets("Merging table files") do |name,xml_filename|
      tables << File.join(output_directory,'intermediate',name,'tables')
    end
    `sort #{tables.map { |t| " '#{t}' "}.join} > #{File.join(output_directory,'intermediate','all_tables')}`
  end
  
  def simplify_worksheets
    worksheets("Simplifying") do |name,xml_filename|
      #fork do
        # i = input( File.join(name,'formulae.ast'))
        # o = output(File.join(name,'missing_functions'))
        # CheckForUnknownFunctions.new.check(i,o)
        # close(i,o)
        simplify_worksheet(name,xml_filename)
      #end
    end
    # missing_function_files = []
    # worksheets("Consolidating any missing functions") do |name,xml_filename|
    #   missing_function_files << File.join(output_directory,'intermediate',name,'missing_functions')
    # end
    # `sort -u #{missing_function_files.map { |t| " '#{t}' "}.join} > #{File.join(output_directory,'intermediate','all_missing_functions')}`
  end
  
  def simplify_worksheet(name,xml_filename)
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
  
  def replace_blanks
    references = {}
    worksheets("Loading formulae") do |name,xml_filename|
      r = references[name] = {}
      i = input(name,"formulae_no_indirects_optimised.ast")
      i.lines do |line|
        ref = line[/^(.*?)\t/,1]
        r[ref] = true
      end
    end
    worksheets("Replacing blanks") do |name,xml_filename|
      #fork do 
        r = ReplaceBlanks.new
        r.references = references
        r.default_sheet_name = name
        replace r, File.join(name,"formulae_no_indirects_optimised.ast"),File.join(name,"formulae_no_blanks.ast")
      #end
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
      
      # Finally, create the output directory
      i = File.join(output_directory,'intermediate',name,"#{basename}#{counter}.ast")
      o = File.join(output_directory,'intermediate',name,finish_filename)
      `cp '#{i}' '#{o}'`
    end
  end
  
  def optimise_sheets(start_filename,finish_filename,basename)
    counter = 1
    
    # Setup start
    worksheets("Setting up for optimise") do |name|
      i = File.join(output_directory,'intermediate',name,start_filename)
      o = File.join(output_directory,'intermediate',name,"#{basename}#{counter}.ast")
      `cp '#{i}' '#{o}'`
    end
    
    worksheets("Replacing with calculated values") do |name,xml_filename|
      #fork do
        replace ReplaceFormulaeWithCalculatedValues, File.join(name,"#{basename}#{counter}.ast"), File.join(name,"#{basename}#{counter+1}.ast")
      #end
    end
    counter += 1
    Process.waitall
      
    references = all_formulae("#{basename}#{counter}.ast")
    inline_ast_decision = lambda do |sheet,cell,references|
      references_to_keep = @values_that_can_be_set_at_runtime[sheet]
      if references_to_keep && (references_to_keep == :all || references_to_keep.include?(cell))
        false
      else
        ast = references[sheet][cell]
        if ast
          if [:number,:string,:blank,:null,:error,:boolean_true,:boolean_false,:sheet_reference,:cell].include?(ast.first)
            #   puts "Inlining #{sheet}.#{cell}: #{ast.inspect}"
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
    
    worksheets("Inlining formulae") do |name,xml_filename|
      #fork do
        r.default_sheet_name = name
        replace r, File.join(name,"#{basename}#{counter}.ast"), File.join(name,"#{basename}#{counter+1}.ast")
      #end
    end
    counter += 1
    Process.waitall
    
    # Finish
    worksheets("Moving sheets") do |name|
      o = File.join(output_directory,'intermediate',name,finish_filename)
      i = File.join(output_directory,'intermediate',name,"#{basename}#{counter}.ast")
      `cp '#{i}' '#{o}'`
    end
  end
  
  def remove_any_cells_not_needed_for_outputs(formula_in = "formulae_no_blanks.ast", formula_out = "formulae_pruned.ast", values_in = "values_no_shared_strings.ast", values_out = "values_pruned.ast")
    if outputs_to_keep && !outputs_to_keep.empty?
      identifier = IdentifyDependencies.new
      identifier.references = all_formulae(formula_in)
      outputs_to_keep.each do |sheet_to_keep,cells_to_keep|
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
        #fork do
          r.cells_to_keep = identifier.dependencies[name]
          rewrite r, File.join(name, formula_in), File.join(name, formula_out)
          rewrite r, File.join(name, values_in), File.join(name, values_out)
        #end
      end
      Process.waitall
    else
      worksheets do |name,xml_filename|
        i = File.join(output_directory,'intermediate',name, formula_in)
        o = File.join(output_directory,'intermediate',name, formula_out)
        `cp '#{i}' '#{o}'`
        i = File.join(output_directory,'intermediate',name, values_in)
        o = File.join(output_directory,'intermediate',name, values_out)
        `cp '#{i}' '#{o}'`
      end
    end
  end
  
  def inline_formulae_that_are_only_used_once
    references = all_formulae("formulae_pruned.ast")
    counter = CountFormulaReferences.new
    count = counter.count(references)
    
    inline_ast_decision = lambda do |sheet,cell,references|
      references_to_keep = @values_that_can_be_set_at_runtime[sheet]
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
      #fork do
        r.default_sheet_name = name
        replace r, File.join(name,"formulae_pruned.ast"), File.join(name,"formulae_inlined.ast")
      #end
    end
    
    remove_any_cells_not_needed_for_outputs("formulae_inlined.ast", "formulae_inlined_pruned.ast", "values_pruned.ast", "values_pruned2.ast")
    
    # worksheets("Skipping inlining") do |name,xml_filename|
    #   i = File.join(output_directory,'intermediate',name, "formulae_pruned.ast")
    #   o = File.join(output_directory,'intermediate',name, "formulae_inlined_pruned.ast")
    #   `cp '#{i}' '#{o}'`
    #   i = File.join(output_directory,'intermediate',name, "values_pruned.ast")
    #   o = File.join(output_directory,'intermediate',name, "values_pruned2.ast")
    #   `cp '#{i}' '#{o}'`
    # end

  end
  
  def separate_formulae_elements
    references = all_formulae("formulae_inlined_pruned.ast")
    identifier = IdentifyRepeatedFormulaElements.new
    repeated_elements = identifier.count(references)
    repeated_elements.each do |sheet,elements|
      elements.delete_if do |element,count|
        count < 2
      end
    end
    worksheets("Writing repeated elements") do |name,xml_filename|
      o = output(name,'common.ast')
      i = 0
      elements = repeated_elements[name]
      elements.each do |element,count|
        o.puts "common#{i}\t#{element}"
        i = i + 1
      end
      close(o)
    end
    worksheets("Replacing repeated elements") do |name,xml_filename|
      replace ReplaceCommonElementsInFormulae, File.join(name,"formulae_inlined_pruned.ast"), File.join(name,"common.ast"), File.join(name,"formulae_inlined_pruned_replaced.ast")
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
  
  def compile_workbook
    compile_workbook_code
    compile_workbook_test
  end
  
  def compile_workbook_code
    w = input("worksheet_ruby_names")
    o = ruby("#{compiled_module_name.downcase}.rb")
    o.puts "# Compiled version of #{excel_file}"
    o.puts "require '#{File.expand_path(File.join(File.dirname(__FILE__),'../excel/excel_functions'))}'"
    o.puts ""
    o.puts "module #{compiled_module_name}"
    o.puts "class Spreadsheet"
    o.puts "  include ExcelFunctions"
    w.lines do |line|
      name, ruby_name = line.strip.split("\t")
      o.puts "  def #{ruby_name}; @#{ruby_name} ||= #{ruby_name.capitalize}.new; end"
    end
    o.puts "end"
    o.puts 'Dir[File.join(File.dirname(__FILE__),"worksheets/","*.rb")].each {|f| autoload(File.basename(f,".rb").capitalize,f)}'
    o.puts "end"
    close(w,o)
  end

  def compile_workbook_test
    w = input("worksheet_ruby_names")
    o = ruby("test_#{compiled_module_name.downcase}.rb")
    o.puts "# All tests for #{excel_file}"
    o.puts  "require 'test/unit'"
    w.lines do |line|
      name, ruby_name = line.strip.split("\t")
      o.puts "require_relative 'tests/test_#{ruby_name.downcase}'"
    end
    close(w,o)
  end
  
  def compile_worksheets
    worksheets("Compiling worksheet") do |name,xml_filename|
      #fork do 
        compile_worksheet_code(name,xml_filename)
        compile_worksheet_test(name,xml_filename)
      #end
    end    
  end
  
  def compile_worksheet_code(name,xml_filename)
    settable_refs = @values_that_can_be_set_at_runtime[name]    
    c = CompileToRuby.new
    c.settable =lambda { |ref| (settable_refs == :all) ? true : settable_refs.include?(ref) } if settable_refs
    i = input(name,"formulae_inlined_pruned_replaced.ast")
    w = input("worksheet_ruby_names")
    ruby_name = ruby_name_for_worksheet_name(name)
    o = ruby('worksheets',"#{ruby_name.downcase}.rb")
    d = output(name,'defaults')
    o.puts "# coding: utf-8"
    o.puts "# #{name}"
    o.puts
    o.puts "require_relative '../#{compiled_module_name.downcase}'"
    o.puts
    o.puts "module #{compiled_module_name}"
    o.puts "class #{ruby_name.capitalize} < Spreadsheet"
    c.rewrite(i,w,o,d)
    o.puts ""
    close(d)
    if settable_refs
      o.puts "  def initialize"
      d = input(name,'defaults')
      d.lines do |line|
        o.puts line
      end
      o.puts "  end"
      o.puts ""
      close(d)
    end
    o.puts "end"
    o.puts "end"
    close(i,o)
  end

  def compile_worksheet_test(name,xml_filename)
    i = input(name,"values_pruned2.ast")
    ruby_name = ruby_name_for_worksheet_name(name)
    o = ruby('tests',"test_#{ruby_name.downcase}.rb")
    o.puts "# coding: utf-8"
    o.puts "# Test for #{name}"
    o.puts  "require 'test/unit'"
    o.puts  "require_relative '../#{compiled_module_name.downcase}'"
    o.puts
    o.puts "module #{compiled_module_name}"
    o.puts "class Test#{ruby_name.capitalize} < Test::Unit::TestCase"
    o.puts "  def worksheet; #{ruby_name.capitalize}.new; end"
    CompileToRubyUnitTest.rewrite(i, o)
    o.puts "end"
    o.puts "end"
    close(i,o)
  end
  
  def ruby_name_for_worksheet_name(name)
    unless @worksheet_names
      w = input("worksheet_ruby_names")
      @worksheet_names = Hash[w.readlines.map { |line| line.split("\t").map { |a| a.strip }}]
      close(w)
    end
    @worksheet_names[name]
  end
    
  def worksheets(message = "Processing",&block)
    IO.readlines(File.join(output_directory,'intermediate','worksheet_names')).each do |line|
      name, filename = *line.split("\t")
      filename = File.expand_path(File.join(xml_dir,'xl',filename.strip))
      puts "#{message} #{name}"
      block.call(name, filename)
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