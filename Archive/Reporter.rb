require_relative 'RDT'
require_relative 'rubyexcel/lib/lib/rubyexcel'
require 'mail' #Send emails automatically

=begin
Example argument hash:
args = {
tool: 'Workshop Reports',
type: 'History Report',
account: 'PHX0%',
ending: get_end_of_last_month,
filename: 'output.txt',
history: false,
key: %w{ CNTX% },
ber: true/false,
starting: get_start_of_last_month,
special: 'PO' / 'PO|L'
title: 'Test Report',
display: false,
email: ['bob@thebuilder.net'],
options: [ 
         [:text_field, :name, 'account', :set, 'PHX0%'],
         [ :checkbox, :id, 'groupbyjob', :set ],
         [ :select_list, :name, 'statuscurrent', :select_value, 'N' ]
         ]
}
=end

#
# This class is now obsolete, replaced by MechReporter
#

class Reporter < RDT

  attr_accessor :args, :data
  
  def initialize( args = { tool: 'Reports' } )
  
    super()
  
    @args = args
    @data = []
    fail ArgumentError, "invalid argument: #{args.inspect}" unless args.kind_of? Hash
  
    temp = %w{username password}
    @user_pass.map!.with_index { |val, idx| puts "Please enter RDT #{ temp[idx] }:"; exit if set_regkey_val( temp[idx], gets.chomp ) == '' } unless @user_pass[0] && @user_pass[1]
  
  end
  
  def []( sym )
    args[ sym ]
  end
  
  def []=( sym, val )
    args[ sym ] = val
  end

  def get_key_from_text_file( filename = '' )
    filename = Dir.pwd.gsub( '/', '\\' ) + '\\list.txt' if filename == ''
    ar = File.read( filename ).split( "\n" ).map( &:strip )
    ar.delete ''
    @args[:key] = ar
  end
  
  def run_report(get_labels=false)
    
    @driver.button( :value => 'Search' ).wait_until_present

    @args[:key] = [ @args[:key] ].flatten
    
    @args[:key].each { |jobref|
 
      if @args[:special] != 'PO|L' || @driver.text_field( :name => 'custordref' ).value != jobref.split('|').first
      
        wait_for_download true if get_labels
        
        case @args[:tool]
        when 'Workshop Reports'
        
          @driver.select_list( :id => 'reptype' ).select @args[:type]
          
          if @args[:special].nil?
            @args[:type] == 'Serial Number Report' ? @driver.text_field( :name => 'serial' ).set( jobref ) : @driver.text_field( :name => 'barcode' ).set( jobref )
          elsif @args[:special] == 'PO'
            @driver.text_field( :name => 'custordref' ).set( jobref )
          elsif @args[:special] == 'PO|L'
            @driver.text_field( :name => 'custordref' ).set( jobref.split('|').first )
          end
          
          @driver.text_field( :name => 'account' ).set @args[:account] if @args[:account]
          @driver.checkbox( :id => 'groupbyjob' ).set( !@args[:history] ) unless @args[:history].nil?
          @args[:ber] ? @driver.select_list( :id => 'berhandle' ).select( 'are always shown' ) : @driver.select_list( :id => 'berhandle' ).select( 'are never shown' )
          
        when 'Reports'
        
          @driver.select_list( :id => 'trackselect' ).select @args[:type]
          @driver.text_field( :id => 'num' ).set jobref if jobref
          @driver.checkbox( :id => 'history' ).set( @args[:history] ) unless @args[:history].nil?
          
        end
        
        if @args[:starting] && @args[:ending]
        
          @driver.select_list(:name => 'befaft').select 'Between'
        
          @driver.select_list(:name => 'dd').select @args[:starting].strftime '%d'
          @driver.select_list(:name => 'mon').select @args[:starting].strftime '%m'
          @driver.select_list(:name => 'yyyy').select @args[:starting].year

          @driver.select_list(:name => 'todd').select @args[:ending].strftime '%d'
          @driver.select_list(:name => 'tomon').select @args[:ending].strftime '%m'
          @driver.select_list(:name => 'toyyyy').select @args[:ending].year
          
        elsif @args[:type] != 'Work In Progress'
        
          @driver.select_list(:id => 'days').select '1 year'
          
        end
        
        set_additional_options unless @args[:options].nil?
        
        @driver.button( :value => 'Search' ).click
        
        save_labels( jobref ) if get_labels
      
      end
      
      @args[:special] == 'PO|L' ? collect_data( jobref.split('|').last ) : collect_data
      
    }
  
    @data.map! { |ar| ar.map! { |el| el.nil? ? el : "#{ el }".gsub( /\s+/, ' ' ).strip } }
      
  end
    
  def set_additional_options
    custom_command @args[:options]
  end
  
  def collect_data( line = nil )
    doc = Nokogiri::HTML @driver.table(:id => 'reporttable').html
    a = doc.css( 'tr' ).map { |row| row.css('th,td').map { |cell| cell.nil? ? nil : cell.text.gsub( /\s+/, ' ' ).strip } }
    a.reject! { |ar| ar.length < 5 }
    a.each { |ar| ar.shift 3 } if @args[:tool] == 'Reports'
    a.select! { |ar| ar[10] =~ /Line Number|\b#{line}\b/i } if line
    a.shift unless @data.empty?
    @data = @data + a unless a.empty?
  end

  def get_cols( headers, multi_array=nil)
    multi_array = @data if multi_array.nil?
    multi_array.transpose.select { |header,_| headers.include?(header) }.sort_by { |header,_| headers.index(header) || headers.length }.transpose  
  end
  
  def save_labels( jobref )
  
    unless @driver.table(:id => 'reporttable').html.include? jobref
      puts "#{jobref} not found"
      return false 
    end
  
    @driver.checkboxes( :name => 'update[]' ).last.set
    
    jobref =~ /STK/i ? @driver.select_list( :name => 'print_link').select( 'stock label' ) : @driver.select_list( :name => 'print_link').select( 'repair label' )
    
    @driver.button( :value => 'Print Checked').click

    @driver.alert.close
    
    wait_for_download

  end
  
  def get_excel
    excel = WIN32OLE::connect( 'excel.application' ) rescue WIN32OLE::new( 'excel.application' )
    excel.visible = true
    excel
  end
  
  def get_workbook( excel=nil )
    excel ||= get_excel
    wb = excel.workbooks.add
    ( ( wb.sheets.count.to_i ) - 1 ).times { |time| wb.sheets(2).delete }
    wb
  end
  
  def dump_to_sheet( data, sheet=nil )
    fail ArgumentError, "Invalid data type: #{ data.class }" unless data.is_a?( Array ) || data.is_a?( RubyExcel )
    data = data.to_a if data.is_a? RubyExcel
    sheet ||= get_workbook.sheets(1)
    sheet.range( sheet.cells( 1, 1 ), sheet.cells( data.length, data[0].length ) ).value = data
    sheet
  end
  
  def make_sheet_pretty( sheet )
    sheet.cells.EntireColumn.AutoFit
    sheet.cells.HorizontalAlignment = -4108
    sheet.cells.VerticalAlignment = -4108
    sheet
  end
  
  def output_data( output = nil, input_array = false, display = nil )
    output ||= 'output.txt'
    display = @args[:display] if display == nil && @args[:display]
    input_array = input_array.to_a if input_array.is_a? RubyExcel::Workbook
    input_array = @data unless input_array
    fail "Not an array: #{ input_array[0].class }" unless input_array[0].class == Array
    output = "#{ Dir.pwd.gsub( '/','\\\\' ) }\\#{ output }" unless output.include? '\\'
    
    #Split between an excel output and a text output
    if output =~ /\.xlsx/i
    
      begin
        sht = dump_to_sheet input_array, ( wb = get_workbook( excel = get_excel ) ).sheets(1)
        excel.visible = false
        excel.DisplayAlerts = false
        sht.name = 'Report'
        make_sheet_pretty sht
      
        wb.saveas output
      
        excel.visible = true if display      
        wb.Close( 0 ) unless display
        excel.Quit() unless excel.visible
      
        output
      
      rescue
      
        excel.DisplayAlerts = true rescue nil
        excel.visible = true rescue nil
        raise
      
      end
    
    elsif output =~ /\.txt/i || output !~ /\./
      output = output + '.txt' if output !~ /\./
      outputstring = ''
      input_array.each { |ar| outputstring << "#{ ar.join("\t") }\n" }
      File.write output, outputstring
      system( 'Notepad.exe', output ) if display
      output
    else
       fail "Unsupported file format: #{ output }"
    end

  end
  
  def output_to_excel_tabs(input_hash = { 'tab_name' => 'Array of Arrays' }, filename = 'ReportOutput.xlsx', display = nil )
    display = @args[:display] if display == nil && @args[:display]
    filename = @args[:filename] if filename == 'ReportOutput.xlsx' && @args[:filename]
    
    fail ArgumentError, "invalid argument: #{input_hash.inspect}" unless input_hash.kind_of? Hash
    fail ArgumentError, "Output must an .xslx file: #{filename.inspect}" if filename.include?( '.' ) && !filename.include?( '.xlsx' )
    input_hash.each { |key, val| fail TypeError, "Input must be: 'Tab name' => 'Array_of_arrays'" unless ( ( key.kind_of?( String ) || key.kind_of?( Symbol ) ) && val[0].kind_of?( Array) ) }
    filename = "#{Dir.pwd.gsub('/','\\\\')}\\#{filename}" unless filename.include? '\\'
    
    begin
      wb = get_workbook( excel = get_excel )
      excel.visible = false
      excel.DisplayAlerts = false
      
      input_hash.each do |name, data|
        
        if wb.sheets(1).name == 'Sheet1'
          sht = wb.sheets(1)
        else
          sht = wb.sheets.add( { 'after' => wb.sheets( wb.sheets.count ) } )
        end
        sht.name = name.to_s.gsub('_',' ')
        make_sheet_pretty( dump_to_sheet( data, sht ) )
        
      end
      
      wb.sheets(1).select
      
      wb.saveas filename
      
      excel.visible = true if display      
      wb.Close(0) unless display
      excel.Quit() unless excel.visible
      
      filename
      
    rescue
    
      excel.DisplayAlerts = true rescue nil
      excel.visible = true rescue nil
      raise
      
    end
    
  end
  
  def run(close_it = true)
    @driver = start false
    login @args[:tool]
    run_report
    close if close_it
    @data
  end
  
  def run_to_file
    run
    output_data @args[:filename]
    @data
  end
  
  def run_email
    run
    output_data @args[:filename], @data, false
    send_email [@args[:filename]], @args[:title], @args[:email]
  end
  
  def run_labels
    get_key_from_text_file
    @driver = start( false, true)
    login @args[:tool]
    run_report true
    close
  end
  
  def void_warranty
    start
    login(@args[:tool])
    tool_address = @driver.url
    until accept_void_warranty
      @driver.goto tool_address
    end
    close
  end
  
  def accept_void_warranty
    doc = Nokogiri::HTML( @driver.table(:id => 'reporttable').html )
    counter = 0
    doc.css('table[@id="reporttable"] tr').each { |row| #loop through all table rows	
      counter +=1
      unless row.css('td[5]').text =~ /warranty/i || row.css('td[10]').text =~ /warranty/i || counter == 1
        cntx, account = row.css('td[6]').text, row.css('td[2]').text
        @driver.goto @driver.table(:id => 'reporttable').td(:text, cntx).parent.a(:index => 1).href
        @driver.text_field(:name,'comments').when_present.set('Automatic Void Warranty Approval.')
        @driver.inputs.each { |inni| inni.click if inni.value == 'Please enter any comments above and press here to confirm agreement' }
        File.open('WarrantyOutput.txt','a') {|warrantyfile| warrantyfile.write("#{cntx}\t#{account}\n") } if account =~ /P[HF][XS]/i
        return false
      end
    }
    true
  end

  def send_email(attachments=[], subject_var=nil, to_ary=nil, body_var='Automated email')
    ( @args[:email].nil? ? to_ary = ( Passwords::MyName.sub(' ', '') + Passwords::EmailSuffix ) : to_ary = @args[:email] ) if to_ary.nil?
    to_ary = [ to_ary ].flatten
    attachments = [ attachments ].flatten if attachments.size != 0
    
    subject_var = 'Automated Email' unless subject_var

    options = { 
      :address               => 'smtp.gmail.com',
      :port                  => 587,
      :user_name             => Passwords::GmailUser,
      :password              => Passwords::GmailPass,
      :authentication        => 'plain',
      :enable_starttls_auto  => true 
     }
    
     Mail.defaults do
      delivery_method :smtp, options
    end

    Mail.deliver do
      from 'Automatic Emailer'
      to to_ary
      subject subject_var
      body body_var
      attachments.each {|filename| add_file(filename) } if attachments.size != 0 && attachments[0]
    end
    
  end
  
end