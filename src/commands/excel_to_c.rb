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
    merge_table_files
    rewrite_worksheets
    simplify_worksheets
    optimise_and_replace_indirect_loop
    replace_blanks
    remove_any_cells_not_needed_for_outputs
    inline_formulae_that_are_only_used_once
    separate_formulae_elements
    compile_workbook
    compile_build_script
    compile_ruby_ffi_interface
  end
  
  def sort_out_output_directories    
    FileUtils.mkdir_p(File.join(output_directory,'intermediate'))
    FileUtils.mkdir_p(File.join(output_directory,'c'))
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
    rewrite MapSheetNamesToCNames, 'worksheet_names', 'worksheet_c_names'
    
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
    o = output('common-elements.ast')
    i = 0
    repeated_elements.each do |element,count|
      o.puts "common#{i}\t#{element}"
      i = i + 1
    end
    close(o)
    
    worksheets("Replacing repeated elements") do |name,xml_filename|
      replace ReplaceCommonElementsInFormulae, File.join(name,"formulae_inlined_pruned_with_sheets.ast"), "common-elements.ast", File.join(name,"formulae_inlined_pruned_replaced.ast")
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
    
    # Probably a better way of getting the runtime file to be compiled with the created file
    puts `cp #{File.join(File.dirname(__FILE__),'..','compile','c','excel_to_c_runtime.c')} #{File.join(output_directory,'c','excel_to_c_runtime.c')}`
    
    
    # Output the workbook preamble
    w = input("worksheet_c_names")
    o = ruby("#{compiled_module_name.downcase}.c")
    o.puts "// Compiled version of #{excel_file}"
    o.puts '#include "excel_to_c_runtime.c"'
    o.puts
    
    # output the common elements
    c = CompileToC.new
    i = input("common-elements.ast")
    c.rewrite(i,w,o)
    close(i)

    # Output the elements from each worksheet in turn
    d = output('defaults')
    worksheets("Compiling worksheet") do |name,xml_filename|
      w.rewind
      settable_refs = @values_that_can_be_set_at_runtime[name]    
      c = CompileToC.new
      c.settable =lambda { |ref| (settable_refs == :all) ? true : settable_refs.include?(ref) } if settable_refs
      c.worksheet = name

      i = input(name,"formulae_inlined_pruned_replaced.ast")
      ruby_name = c_name_for_worksheet_name(name)
      o.puts "// start #{name}"
      c.rewrite(i,w,o,d)
      o.puts "// end #{name}"
      o.puts
      close(i)
    end
    close(d)
    
    # Output a function that can set settable values to their defaults
    o.puts "void set_to_default_values() {"
    d = input('defaults')
    d.each_line do |line|
      o.puts line
    end
    o.puts "}"
    o.puts
    
    close(w,o,d)
  end
  
  def compile_build_script
    o = ruby("Makefile")
    name = compiled_module_name.downcase
    
    # Target for shared library
    o.puts "lib#{name}.dylib: #{name}.o"
    o.puts "\tgcc -shared -o lib#{name}.dylib #{name}.o"
    o.puts
    
    # Target for compiled version
    o.puts "#{name}.o:"
    o.puts "\tgcc -Wall -fPIC -c #{name}.c"
    o.puts
    
    # Target for cleaning
    o.puts "clean:"
    o.puts "\trm #{name}.o"
    o.puts "\trm lib#{name}.dylib"
    
    close(o)
  end
  
  def compile_ruby_ffi_interface
    all_formulae = all_formulae('formulae_inlined_pruned_replaced.ast')
    name = compiled_module_name.downcase
    o = ruby("#{name}.rb")
    code = <<END
require 'ffi'

module #{name.capitalize}
  extend FFI::Library
  ffi_lib '#{name}'
  ExcelType = enum :ExcelEmpty, :ExcelNumber, :ExcelString, :ExcelBoolean, :ExcelError, :ExcelRange
                
  class ExcelValue < FFI::Struct
    layout :type, ExcelType,
  	       :number, :double,
  	       :string, :string,
         	 :array, :pointer,
           :rows, :int,
           :columns, :int             
  end
  
END
    o.puts code
    o.puts "  # use this function to set all cells to the values they had when the sheet was compiled"
    o.puts "  attach_function 'set_to_default_values', [], :void"
    
    worksheets("Adding references to ruby shim for") do |name,xml_filename|
      o.puts
      o.puts "  # start of #{name}"  
      c_name = c_name_for_worksheet_name(name)

      # Put in place the setters, if any
      settable_refs = @values_that_can_be_set_at_runtime[name]
      if settable_refs
        settable_refs = all_formulae[name].keys if settable_refs == :all
        settable_refs.each do |ref|
          setter = "set_#{c_name}_#{ref.downcase}"
          type = case all_formulae[name][ref.upcase].first
          when :number, :percentage
            ':double'
          when :error
            ":int"
          when :string
            ":string"
          when :boolean_true, :boolean_false
            ":int"
          else
            raise NotSupportedException.new("#{all_formulae[name][ref.downcase]} can't be settable")
          end
          o.puts "  attach_function '#{setter}', [#{type}], :void"
        end
      end

      if !outputs_to_keep || outputs_to_keep.empty? || outputs_to_keep[name] == :all
        getable_refs = all_formulae[name].keys
      elsif !outputs_to_keep[name] && settable_refs
        getable_refs = settable_refs
      else
        getable_refs = outputs_to_keep[name] || []
      end
        
      getable_refs.each do |ref|
        o.puts "  attach_function '#{c_name}_#{ref.downcase}', [], ExcelValue.by_value"
      end
        
      o.puts "  # end of #{name}"
    end
    o.puts "end"  
    close(o)
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
    File.open(File.join(output_directory,'c',*args),'w')
  end
  
  def close(*args)
    args.map(&:close)
  end
  
end