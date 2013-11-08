require_relative 'RDT'
require_relative 'WinBoxes'
require_relative 'MechReporter'

class Despatch < RDT
  include WinBoxes

  attr_accessor  :data, :errors, :test, :copies

  def initialize( test = false )
    
    ocra_build if defined?(Ocra)
  
    super(test, true)
    @test = test
    @data = Hash.new { |hash, key| hash[ key ] = [] }
    @errors = []
    @copies = 1
    
  end

  def await_input( printout = true )
    
    printout = false if @test
    
    prompt = 'Scan the first CNTX Number:'
    prev = 'CNTX'
    loop do
    
      result = inputbox( prompt, prev )
      
      if result.nil? && @data.empty?
        @driver.close rescue nil
        raise StandardError, 'User Quit'
      end
      
      break if result.nil?
      next if result.empty?
      prompt = "Scan the next CNTX Number |OR| Press Cancel to Start Despatching."
      
      begin
        result = cntxify( result.upcase.strip )
      rescue ArgumentError
        msgbox "Invalid CNTX number: #{ result }", 'Error', VBCRITICAL + VBSYSTEMMODAL
        next
      end
      
      @data['Input'] << result unless @data['Input'].include?( result )
      prev = 'Scanned: ' + result
      
    end
    
    if printout
      @copies = ( inputbox( 'Enter the number of copies of each Despatch Note:', 'Print Copies', 1 ).strip.to_i rescue 1 )
    else
      @copies = 1
    end
    
  end

  def cntxify( str )
    return str if str.length == 14
    num = str.to_i
    raise ArgumentError, "#{ str } is not a valid CNTX number!" if num.zero?
    'CNTX' + '%010d' % num
  end
  
  def check_page
    @driver ||= start( true, true, true )
    @driver.window(:url => /pid=873/ ).use rescue login( 'Despatch Tool' )
  end

  def despatch_all
    @data.each do | account, jobs |
    
      # Search by Account
      until @driver.hidden(:id => 'value_cust' ).value =~ /\A#{account}\z/i
        @driver.text_field( :id => 'cust' ).when_present.set( account )
        sleep 0.5
        @driver.text_field( :id => 'cust' ).send_keys( :tab )
        sleep 2
      end
      @driver.text_field( :id => 'cust' ).when_present.flash
      @driver.table(:index => 1 ).wait_until_present
      
      ticked = false
      
      # Tick the boxes
      jobs.each do |ref|
        begin
          @driver.checkbox( :id => ref ).set
          @driver.table(:index => 1 ).row( :id => 'row_' + ref ).td(:index => 0).flash
          ticked = true
        rescue Watir::Exception::UnknownObjectException
          @errors << ref + ' is missing from the Despatch Tool'
        end
      end
      
      # Generate Despatch Note
      if ticked
        wait_for_download true
        @driver.button( :name => 'action' ).click
        print_label wait_for_download, @copies
      end
      
    end
  end
  
  def get_accounts
  
    # Set up the Accounts Report
    uri = RDTQuery.new( 'http://' + ( @test ? Password::RDTTestDomain : Password::RDTCoreDomain ) + '/redetrack/bin/report.php?report_locations_list=T&report_limit_to_top_locations=N&action=boxtrack&num=CNTX0000998877&pod=A&tf=current&days=365&befaft=b&dd=19&mon=07&yyyy=2013&timetype=any&fdays=1&fbefaft=a&fdd=19&fmon=07&fyyyy=2013&ardd=19&armon=07&aryyyy=2012' )
    uri.set_dates( Date.today )
    
    # Run the Accounts Report
    m = MechReporter.new( true, @test )
    m.user, m.pass = @user_pass
    result = m.run_keyed( uri, @data[ 'Input' ] ).gc( 'Account', 'Bar Code' ).uniq( 'Bar Code' ).sort_by( 'Account' )
    fail NoMethodError, 'No Accounts Found' if result.maxrow < 2
    
    result.rows(2) do |r|
      @data[ r.val( 'Account' ).upcase ] << r.val( 'Bar Code' )
      @data['Input'].delete( r.val( 'Bar Code' ) )
    end
    
    # Cleanup
    @data[ 'Input' ].each { |ref| @errors << ref + ' not found' }
    @data.delete( 'Input' )
    @data[ '' ].each { |ref| @errors << ref + ' has a blank Account Number' }
    @data.delete( '' )
    
  end
  
  def ocra_build
    Crypt::Blowfish.new('1').encrypt_string('Moose')
    Watir::Browser.new.close
    MechReporter.new
  end
  
  def report_errors
    return false unless @errors.any?
    msg = $/ + Time.now.to_s + $/ + 'Completed with Errors:' + $/ + @errors.join( $/ ) + $/
    errbox( msg )
    File.open( ENV['HOME'] + '/Desktop/Despatch Errors.log', 'a' ) { |f| f.write msg }
  end
  
  def run
    
    printer = Win32::Registry::HKEY_CURRENT_USER.open('Software\Microsoft\Windows NT\CurrentVersion\Windows')['Device'].split(',').first
    fail NameError, 'Invalid Printer: ' + printer if printer =~ /zebra|lp28|label|ZDesigner/i && !@test
  
    test_login
  
    loop do
  
      await_input
      
      get_accounts
      
      check_page
      
      despatch_all
      
      report_errors
      
      @data.clear
      @errors.clear
      
    end
    
  rescue => error
    if error.to_s == 'User Quit'
      @driver.close rescue nil
    else
      errbox( error.inspect + $/ + error.backtrace.to_s )
    end
  end 
  
  def test_login
    m = MechReporter.new( true, @test )
    m.user, m.pass = @user_pass
    m.login
  end

end
