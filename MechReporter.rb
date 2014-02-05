require 'rubygems'                             # Try and fix the "no such file to load" errors
require 'addressable/uri'                      # Encoding
require 'csv'                                  # CSV Parsing
require_relative 'DateTools'                   # Date handling
require 'mail'                                 # Send emails automatically
require 'mechanize'                            # Browser
require 'nokogiri'                             # Explicit require for OCRA
require_relative 'rubyexcel/lib/lib/rubyexcel' # My shiny data handling gem
require 'timeout'                              # Prevent it waiting for eternity
require_relative 'RegistryTools'               # Registry
require_relative 'Holidays'                    # List of Bank Holiday dates: BankHolidays.all
require_relative 'Passwords'                   # Passwords & personal info

#
# Runs ReDeTrack reports quickly and invisibly
#
# @note Use "require_relative 'mechreporter/gmailer'" to load #send_gmail method.
#

class MechReporter
  include DateTools, RegistryTools
  
  # Aim label files at the Documents folder
  LabelFolder = RubyExcel.documents_path
  # Permit access to the Mechanize agent for debugging
  attr_accessor :agent
  
  # Get and set the RDT username
  attr_accessor :user
 
  # Get and set the RDT password
  attr_accessor :pass
  
  # Turn console output on / off
  attr_accessor :verbose
  
  # PHP session ID
  attr_accessor :sess
  
  # Domain
  attr_accessor :domain
  
  #
  # Create an instance of MechReporter
  #
  
  def initialize( verbose = true, test = false )
  
    # Allow EXE build without running.
    if defined?(Ocra)
      # Let OCRA pick up the win32ole extension
      RubyExcel::Workbook.new.documents_path
      # And the cookies from Mechanize
      Mechanize.new.cookies
      exit
    end
  
    # Get the username and password from the registry
    self.user, self.pass = get_regkey_val('username'), get_regkey_val('password')
    
    # Ask for them if they're not found
    unless user && pass
      self.user, self.pass = [ 'username', 'password' ].map { |name|
        puts "Please enter RDT #{ name.capitalize }:"
        exit if ( res = gets.chomp ).empty?
        set_regkey_val( name, res )
      }
    end
    
    # Create a Mechanize agent to do the website navigation
    self.agent = Mechanize.new do |agent|
      agent.user_agent_alias = 'Windows Mozilla'
      agent.agent.http.verify_mode = OpenSSL::SSL::VERIFY_NONE
      agent.open_timeout = 120
      agent.read_timeout = 120
    end
    
    self.sess = ''
    self.verbose = verbose
    self.domain = test ? Passwords::RDTTestDomain : Passwords::RDTCoreDomain
  end
  
  #
  # Current Acrobat Reader Executable
  #
  # @return [String] the full file path
  #
  
  def acrobat
    Win32::Registry::HKEY_LOCAL_MACHINE.open('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe')['']
  end
  
  #
  # Convert input to a valid CNTX format
  #
  # @param [String] str the string to convert to a CNTX number
  # @return [String]
  #
  
  def self.cntxify( str )
    return str if str.length == 14
    num = str.to_i
    raise ArgumentError, "#{ str } is not a valid CNTX number!" if num.zero?
    'CNTX' + '%010d' % num
  end
  
  # {MechReporter.cntxify}
  
  def cntxify( str )
    self.class.cntxify( str )
  end
  
  #
  # Debugging tool to write the current page's html to a file
  #
  
  def dump_page( output_file = ENV['HOME'] + '/Desktop/a.htm' )
    File.write( output_file, agent.page.body )
  end

  #
  # Common Report: Return all Customers as a Hash (uppercase keys)
  #
  # @return [Hash]
  #
  
  def get_customers
   
    # Go to the Customers tool
    goto_tool 'redetrack/bin/manager'
    
    # Extract the table as a RubyExcel::Sheet
    s = RubyExcel::Workbook.new.load( Nokogiri::HTML( agent.page.body ).css( 'tr' ).map { |row| row.css('th,td').map { |cell| cell.nil? ? nil : cell.text.gsub( /\s+/, ' ' ).strip } } )
    
    # Remove the buttons and blank rows
    s.row(2).delete
    s.compact!
    
    # Store the account details in a Hash by :name and :customer
    h = Hash.new({})
    s.rows(2) do |row|
      name = ( row['D'].empty? ? 'Other' : row['D'] )
      customer = ( row['E'].empty? ? 'Other' : row['E'] )
      h[ row['C'].upcase ] = { name: name, customer: customer }
    end
    h
  end
  
  #
  # Common Report: Return all Locations as a Hash, with the value as "FSL" or "MK"
  #
  # @return [Hash]
  #
  
  def get_locations( fsl = true )
    loc = run( 'http://' + domain + '/redetrack/bin/mylocations.php?pid=249&rnd=3865&uid=2366&org=33&rdt_site_id=222&si=243')
    loc.column(1).delete
    loc.range('D:G').delete
    loc.insert_rows( 1, 1 )
    loc['A1:C1'] = [ 'Name', 'City', 'Parent Location' ]
    hash = {}
    loc.rows(2) do |row|
      hash[ row['A'].upcase.strip ] = if row['B'].nil? || row['B'] !~ /Milton Keynes/i 
        fsl ? 'FSL' : ( row['B'].nil? ? 'Blank' : row['B'] )
      else 
        'MILTON KEYNES'
      end
    end
    hash
  end
  
  #
  # Go to a specific RDT tool
  #
  # @param [String] tool_uri the section of the tool's link which ends with '.php'
  #

  def goto_tool( tool_uri='report' )
    login
    agent.get( 'http://' + domain + '/' + agent.page.links.find { |l| l.attributes[:onclick] =~ /#{ tool_uri }\.php/ }.attributes[:onclick].split("'")[1] )
  end
  
  #
  # Allows verbose setting to catch console output
  #
  
  def internal_puts( *msg )
    puts *msg if verbose
  end
  
  #
  # Login to RDT
  #
  
  def login

    #Go to login page
    agent.get( 'http://' + domain )
  
    # Don't log in twice
    return agent if agent.page.body.include?( 'logged in' )
    
    # Output status
    internal_puts 'logging in...'
    
    # Fill in and submit the RDT login form
    f = agent.page.forms.first
    f.username, f.password, f.org = user, pass, Passwords::OrgID
    agent.submit(f, f.buttons.first)
    
    # Wait for a response
    wait_for_it( 3, LoginFailed.new( 'Failed to log into RDT.' ) ) { agent.page.body.include?('logged in') }
    
    # Report success
    internal_puts 'logged in: ' + user
    
    # Extract session ID
    self.sess = agent.cookies.find { |c| c.name == 'PHPSESSID' && c.domain == domain }.value
    
    # Return self for the next method in the chain
    self
  end

  #
  # Default Printer Name
  #
  # @return [String] the name of the default printer
  #
  
  def printer
    Win32::Registry::HKEY_CURRENT_USER.open('Software\Microsoft\Windows NT\CurrentVersion\Windows')['Device'].split(',').first
  end
  
  #
  # Print a pdf file to the default printer
  #
  # @param [String] file_name the full path of the pdf file
  # @param [Fixnum] copies the number of copies to print
  #
  
  def print_label ( file_name, copies=1 )
    internal_puts 'Printing Labels...'
    cmd = %Q|"#{acrobat}" /n /s /h /t "#{ file_name.gsub('/', '\\') }" "#{ printer }"|.gsub('\\','\\\\')
    copies.times do
      sleep 0.3
      Thread.new { system cmd }
    end
  end
  alias print_labels print_label
  
  #
  # Display an Adobe Print Prompt for a pdf file
  #
  # @param [String] file_name the full path of the pdf file
  #
  
  def print_prompt ( file_name )
    internal_puts 'Opening Print Dialog...'
    Thread.new { `"#{acrobat}" /p "#{ file_name.gsub('/', '\\') }"` }
  end
  
  #
  # Download and save a set of labels from RDT
  #
  # @param [Array<String>] cntx_numbers the job references to print
  # @return [String] the filename
  #
  
  def save_labels( cntx_numbers, type = 1 )
  
    login
  
    link = case type
    
    when 1
      '/redetrack/bin/centrexlabels.php?bc='
    when 2
      '/redetrack/bin/centrexstocklabels.php?bc='
    when 3
      '/redetrack/bin/centrexlabels.php?label=repair&bc='
    else
      raise ArgumentError, "Invalid label type: #{ type }. Must be 1, 2, or 3 "
    end
  
    internal_puts 'Downloading Labels...'
    file = agent.get( link + cntx_numbers.join(',') )
    filename = LabelFolder + '\\AutoLabels.pdf'
    file.save! filename
    filename
    
  end
  
  #
  # Extract a CSV Report from RDT
  #
  # @param [String] uri the URI to run the report with
  # @return [RubyExcel::Sheet] the CSV data as a RubyExcel Sheet
  #
  
  def run( uri, test = false )
    
    login
    internal_puts 'running report...'
    
    # Convert the report to only return a csv
    uri = RDTQuery.new( uri )
    uri[ 'givecsvonly' ] = 'Yes'
    
    # Get the report page from our list of arguments
    page = agent.get( uri.to_s )
    
    # Wait until the session ID shows up with our CSV link (AJAX)
    regex = /[\w\/]+#@sess[\w\/\.]+/
    internal_puts 'waiting for download link'
    wait_for_it( 60, MechReporterTimeout.new( 'Failed to find CSV download link.' ) ) { page.body =~ regex }
    filename = page.body.match( regex ).to_s
    
    internal_puts 'downloading report file...'
    
    # Download & Interpret the CSV, and drop the data into RubyExcel
    RubyExcel::Workbook.new.load( CSV.parse( agent.get('http://' + domain + filename ).content ) )
  end
  
  # 
  # Runs the report with a given set of keys.
  #   Breaks long queries into smaller chunks.
  #
  # @param [String, RDTQuery] uri the query to run
  # @param [String] keyname the name of the query value to use as a key
  # @param [Array<String>] keys the Array of keys to run the report with
  # @return [RubyExcel::Sheet]
  #
  
  def run_keyed( uri, keys, keyname = 'num', test = false )
  
    # Avoid modifying the variable passed into the method (and avoid blanks)
    keys = keys.compact
  
    # Set the upper byte limit for the URI
    maxlen = 7000
    
    # Set up the current query
    current_query = RDTQuery.new uri
    
    # Wipe existing numbers from the URI
    current_query[ keyname ] = ''
    
    # Determine the maximum references we can report with
    allowed_extra = maxlen - URI.escape( current_query.to_s ).bytesize
    
    # Create an empty sheet to populate with data
    res = RubyExcel::Workbook.new.add
    
    # Keep going until we've gathered all the requisite data
    until keys.empty?
    
      # Wipe out the query key for each loop
      stringy = ''
    
      # Report progress back to the impatient user
      internal_puts 'Remaining: ' + keys.length.to_s
      
      # Keep building the string until we hit the limit
      until keys.empty? || URI.escape( stringy ).bytesize + URI.escape( keys.last ).bytesize >= allowed_extra
    
        # Take from array, add to string
        stringy << keys.pop << '|'
      
      end
      
      # Build the finished URI
      current_query[ keyname ] = stringy.chomp('|')
      
      # Run the query and return the data
      res << run( current_query, test )
      
    end # keys.empty
    
    # Return the Sheet
    res
  end
  
  
  # 
  # Runs the report with a given set of keys.
  #   Use this when pipe character seperation support is unavailable.
  #
  # @param [String, RDTQuery] uri the query to run
  # @param [String] keyname the name of the query value to use as a key
  # @param [Array<String>] keys the Array of keys to run the report with
  # @return [RubyExcel::Sheet]
  #
  
  def run_keyed_individual( uri, keys, keyname = 'num' )
    
    # Set up the current query
    current_query = RDTQuery.new uri
    
    # Wipe existing numbers from the URI
    current_query[ keyname ] = ''
    
    # Create an empty sheet to populate with data
    res = RubyExcel::Workbook.new.add
    
    keys.each do |key|
    
      internal_puts key
    
      # Update the URI
      current_query[ keyname ] = key

      # Run the query and return the data
      res << run( current_query )
      
    end
    
    # Return the Sheet
    res
  end
  
  #
  # Run a Management Report with the given options
  #
  
  def run_mr( by: 'Months', length: 1, type: 'Customer Demand', split_by: 'Account', then_by: nil, select_by: nil, select_val: nil )
    
    login
    internal_puts 'running report...'
    
    goto_tool 'mr/bin/main'
    
    f = agent.page.forms.first
    [ [ :tb_id, by ], [ :tb_n, length ], [ :m_id, type ], [ :b_id, split_by ], [ :b_id_2, then_by ], [ :s_id, select_by ], [ :s_val, select_val ] ].each do |name, val|
      
      if val
      
        field = f.field( name.to_s )
        if name == :tb_n
          field.value = val
        elsif field.respond_to?( :options )
          field.option_with( :text => val.to_s ).select
        else
          field.value = val
        end
        
      end
      
    end
    agent.submit(f, f.buttons.first)
    
    # Wait until the session ID shows up with our CSV link (AJAX)
    regex = /[\w\/]+#@sess[\w\/\.]+/
    internal_puts 'waiting for download link'
    wait_for_it( 60, MechReporterTimeout.new( 'Failed to find CSV download link.' ) ) { agent.page.body =~ regex }
    filename = agent.page.body.match( regex ).to_s
    
    internal_puts 'downloading report file...'
    
    # Download & Interpret the CSV, and drop the data into RubyExcel
    RubyExcel::Workbook.new.load( CSV.parse( agent.get('http://' + domain + filename ).content ) )

  end
  
  #
  # Break a report into smaller time-frames in order to avoid timeouts
  #
  # @param [String] uri the URI to run the report with
  # @param [Date] from the oldest Date required
  # @param [Date] to the newest Date required
  # @param [Fixnum] step the maximum number of days to run at once
  # @return [RubyExcel::Sheet] the CSV data as a RubyExcel Sheet
  #
  
  def run_staggered( uri, from, to, step )
  
    # Create a query
    query = RDTQuery.new( uri )
    
    # Create a silent version of MechReporter
    m = MechReporter.new( false )
    
    # Create a blank Sheet
    ret = RubyExcel::Workbook.new.add 'Report'
    
    # Run the report, stepping through the dates until the end
    ( from..to ).each_slice( step ) { |range| puts "Running date range #{ range.first.strftime( '%d/%m/%y' ) } - #{ range.last.strftime( '%d/%m/%y' ) }"; ret << m.run( query.set_dates( range.first, range.last ) ); internal_puts ret.maxrow.to_s + ' lines' }
    
    # Return the combined data
    ret
    
  end
  
  def run_staggered_with_retry( uri, from, to, step )
  
    # Create a query
    query = RDTQuery.new( uri )
    
    # Create a silent version of MechReporter
    m = MechReporter.new( false )
    
    # Create a blank Sheet
    ret = RubyExcel::Workbook.new.add 'Report'
    
      # Run the report, stepping through the dates until the end
      ( from..to ).each_slice( step ) { |range|
        begin
          puts "Running date range #{ range.first.strftime( '%d/%m/%y' ) } - #{ range.last.strftime( '%d/%m/%y' ) }"
          ret << m.run( query.set_dates( range.first, range.last ) ) rescue retry
        rescue => err
          puts err, 'retrying'
          retry
        end
        internal_puts ret.maxrow.to_s + ' lines'
      }
    
    # Return the combined data
    ret
    
  end
  
  #
  # Break a report into threads to run it in parallel
  #
  # @param [String] uri the URI to run the report with
  # @param [Date] from the oldest Date required
  # @param [Date] to the newest Date required
  # @param [Fixnum] step the maximum number of days to run at once
  # @return [RubyExcel::Sheet] the CSV data as a RubyExcel Sheet
  #
  
  def run_threaded( uri, from, to, step )
  
    # Create a query
    query = RDTQuery.new( uri )
    
    # Create a blank Sheet
    ret = RubyExcel::Workbook.new.add 'Report'
    
    # Run the report, stepping through the dates until the end
    ( from..to ).each_slice( step ).map do |range|
      Thread.new do
        internal_puts "Running date range #{ range.first.strftime( '%d/%m/%y' ) } - #{ range.last.strftime( '%d/%m/%y' ) }"
        m = MechReporter.new
        m.verbose = false
        Thread.current[:output] = m.run( query.set_dates( range.first, range.last ) )
        internal_puts "Completed date range #{ range.first.strftime( '%d/%m/%y' ) } - #{ range.last.strftime( '%d/%m/%y' ) }"
      end
    end.each { |thread| thread.join; ret << thread[:output] }
    
    # Return the combined data
    ret
    
  end

  #
  # Send an email through an email account
  #
  # @param [String, Array<String>] attachments the filename(s) to attach
  # @param [String] subject_var the email subject
  # @param [String, Array<String>] to_ary the "To" email address(es)
  # @param [String] body_var the email body
  #
  
  def send_email( attachments=[], subject_var='Automated Email', to_ary=Passwords::MyName, body_var='Automated email', bcc_ary = [] )

    # Standardise email addresses into an array in case a string was passed through
    to_ary = [ to_ary ].flatten.compact
    bcc_ary = [ bcc_ary ].flatten.compact
    
    # Remove whitespace and add a domain if required.
    # This allows more readable and DRY names to be passed into the method.
    mailify = -> name { name.include?('@') ? name : ( name + Passwords::EmailSuffix ).gsub(/\s/,'') }
    to_ary.map! { |name| mailify[ name ] }
    
    bcc_ary = bcc_ary.map { |name| mailify[ name ] }
    
    # Make sure attachments is an array
    attachments = [ attachments ].flatten if attachments.size != 0

    options = { address:              Passwords::EmailServer,
                port:                 25,
                user_name:            Passwords::EmailUser,
                password:             Passwords::EmailPass,
                authentication:       'plain',
                enable_starttls_auto: true }
    
    Mail.defaults do
      delivery_method :smtp, options
    end

    Mail.deliver do
        from Passwords::EmailUser + Passwords::EmailSuffix
        to to_ary
        subject subject_var
        attachments.each { |filename| add_file(filename) } unless attachments.empty?
        body body_var
        bcc bcc_ary
    end

  end
  
  #
  # Custom wait method
  #
  # @param [Fixnum] seconds the number of seconds to wait
  #
  
  def wait_for_it( seconds = 60, failure = MechReporterTimeout.new( 'Timed out' ) )
    
    timer = 0
    
    # Failsafe in case it all goes wrong
    Timeout::timeout(seconds) do
      
      # Until whatever argument is passed is true
      until yield
        
        # Show a counting timer
        if verbose
          string = 'Waiting: ' + timer.to_s + ' second(s) '
          print string
        end
        sleep 1; timer+=1
        string.length.times { print "\b" } if verbose
        
      end
      
    end
    
    # Make sure we leave the console on a new line
    print "\n" if timer > 0 && verbose
    
  rescue Timeout::Error
  
    # On error, send a newline to STDOUT and raise to the next error handler
    print "\n" if verbose
    raise failure
    
  end

  #
  # Display the class as a String
  #
  # @return [String]
  #
  
  def to_s
    self.class.to_s + ' - ' + user.to_s
  end

end

#
# Handles URIs
#

class RDTQuery
  include DateTools

  # Provide access to the URI object
  attr_reader :query

  # Default WIP report URI
  WIP = 'http://' + Passwords::RDTCoreDomain + '/redetrack/bin/centrexticketreport.php?reptype=wip2&groupbyjob=Yes&berhandle=N&status=-1&statuscurrent=N&days=1&nolimit=0&range=nolimit&depot=Centrex+Computing+Services&action=ticketreport&go=Search'
  
  # Default History report URI
  HIST = 'http://' + Passwords::RDTCoreDomain + '/redetrack/bin/centrexticketreport.php?reptype=hist&item_type=&barcode=&engineer=&serial=&custordref=&groupbyjob=Yes&berhandle=Y&account=&clientnum=&status=-1&statuscurrent=N&nolimit=0&range=daterange&befaft=f&dd=28&mon=04&yyyy=2013&todd=28&tomon=04&toyyyy=2013&depot=Centrex+Computing+Services&action=ticketreport&go=Search'

  # Default Barcode Report - Latest - URI
  BCL = 'http://' + Passwords::RDTCoreDomain + '/redetrack/bin/report.php?report_locations_list=T&select_div_last_shown=&report_limit_to_top_locations=N&action=boxtrack&num=&num_raisedtrack=&status=&pod=A&status_code=&itemtype=&location=&value_location=&tf=current&days=365&befaft=b&dd=14&mon=10&yyyy=2013&timetype=any&fdays=1&fbefaft=a&fdd=14&fmon=10&fyyyy=2013&ardd=14&armon=10&aryyyy=2012'
  
  # Default Barcode Report - Full - URI
  BCF = 'http://' + Passwords::RDTCoreDomain + '/redetrack/bin/report.php?report_locations_list=T&select_div_last_shown=&report_limit_to_top_locations=N&action=boxtrack&num=&num_raisedtrack=&status=&pod=A&status_code=&itemtype=&location=&value_location=&tf=current&days=365&befaft=b&dd=14&mon=10&yyyy=2013&fdays=1&fbefaft=a&fdd=14&fmon=10&fyyyy=2013&ardd=14&armon=10&aryyyy=2012&history=on'
  
  # Default Workshop Report - Latest - URI
  WSL = 'http://' + Passwords::RDTCoreDomain + '/redetrack/bin/centrexticketreport.php?reptype=job&item_type=&barcode=&engineer=&serial=&custordref=&includeuaj=Yes&groupbyjob=Yes&berhandle=Y&account=&clientnum=&status=-1&statuscurrent=N&days=365&nolimit=0&range=daterange&befaft=b&dd=14&mon=10&yyyy=2013&depot=Centrex+Computing+Services&action=ticketreport&go=Search'

  # Default Workshop Report - Full - URI
  WSF = 'http://' + Passwords::RDTCoreDomain + '/redetrack/bin/centrexticketreport.php?reptype=job&item_type=&barcode=&engineer=&serial=&custordref=&includeuaj=Yes&berhandle=Y&account=&clientnum=&status=-1&statuscurrent=N&days=365&nolimit=0&range=daterange&befaft=b&dd=14&mon=10&yyyy=2013&depot=Centrex+Computing+Services&action=ticketreport&go=Search'
  
  #
  # Create a new instance of RDTQuery
  #
  # @param [String] string the URI to parse
  #
  
  def initialize( string = '' )
    self.query = string.to_s
  end

  #
  # Get a value from a query
  #
  # @param [String, Symbol] ref the name of the value to return
  #
  
  def []( ref )
    query.query_values[ ref.to_s ]
  end
  
  #
  # Update a value in the query
  #
  # @param [String, Symbol] ref the name of the value to update
  # @param [String] val the value to add to the query
  #
  
  def []=( ref, val )
    update_query( ref => val )
    val
  end
  
  #
  # Set a new URI
  # 
  # @param [String] uri the URI to use
  #
  
  def query=( uri )
    @query = Addressable::URI.parse( uri.to_s )
    #self.query = URI.parse( uri.to_s )
    uri
  end
  
  #
  # Set the query dates to the given date(s)
  #
  # @param [Date] start_date the date to start the search at
  # @param [Date] end_date the date to end the search at
  #
  
  def set_dates( start_date, end_date = nil )
  
    # If the end date isn't specified, default to sameday
    end_date ||= start_date
    
    # Break the dates into a hash query format. 
    hash = [ [ start_date, { dd: '%d', mon: '%m', yyyy: '%Y' } ] , [ end_date, { todd: '%d', tomon: '%m', toyyyy: '%Y' } ] ].inject({}) do |h, ( date, args )|
      args.each { |k,v| h[k.to_s] = date.strftime(v) } ; h
    end

    # Update the query with the new values
    update_query( hash )
  end
  
  #
  # Set the query date and time to the given datetimes
  #
  # @param [Time] start_datetime the datetime to start the search at
  # @param [Time] end_datetime the datetime to end the search at
  #
  
  def set_datetimes( start_datetime, end_datetime )
  
    start_datetime, end_datetime = [ start_datetime, end_datetime ].sort
  
    set_dates( start_datetime.to_date, end_datetime.to_date )
  
    # Break the dates into a hash query format. 
    hash = [ [ start_datetime, { fthours: '%H', ftmins: '%M' } ] , [ end_datetime, { tthours: '%H', ttmins: '%M' } ] ].inject({}) do |h, ( date, args )|
      args.each { |k,v| h[k.to_s] = date.strftime(v) } ; h
    end
    
    hash[ 'timetype' ] = 'between'
    hash[ 'befaft' ] = 'f'
    
    # Update the query with the new values
    update_query( hash )
  end
  
  #
  # Make changes to an existing query, only overwriting duplicate keys
  #
  # @param [Hash] hash the Hash of query options to set
  #
  
  def update_query( hash )
  
    # Convert symbols to strings to avoid duplicate entries
    hash = Hash[hash.map {|k, v| [ k.to_s, v] }] if hash.keys.any? { |k| k.is_a?(Symbol) }
    
    # Merge the changes with the existing query
    query.query_values = query.query_values.merge!( hash )
    
    self
  end
  
  #
  # View the object for debugging
  #
  # @return [String]
  # 
  
  def inspect
    "<#{ self.class }:0x#{ '%x' % (object_id << 1) } - #{ query.basename.sub( /\..+/, '' ) }>"
  end
  
  #
  # Retreive the URI as a string
  #
  # @return [String]
  #
  
  def to_s
    query.to_s
  end
  
end

class LoginFailed < StandardError
end

class MechReporterTimeout < StandardError
end