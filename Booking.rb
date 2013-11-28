require_relative 'RDT'
require_relative 'WinBoxes'
require_relative 'MechReporter'

class Booking < RDT
  include WinBoxes

  attr_accessor  :data, :warranty, :error, :preadvice, :excel

  def initialize( test = false )
    super(test)
    self.test = test
    self.data = {}
    temp = %w{username password}
    
    unless user_pass[0] && user_pass[1]
      user_pass.map!.with_index { |val, idx|
        puts "Please enter RDT #{ temp[idx] }:"
        input = set_regkey_val( temp[idx], gets.chomp )
        exit if input == ''
        input
      }
    end

  end

  def cntxify( str )
    return str if str.length == 14
    num = str.to_i
    raise ArgumentError, "#{ str } is not a valid CNTX number!" if num.zero?
    'CNTX' + '%010d' % num
  end
  
  def get_excel_input
  
    preadvice, part_list = nil, nil
    excel = WIN32OLE::connect('excel.application')
    self.excel = excel
    
    preadvice = excel.activesheet if excel.activesheet.range('A1').value =~ /cenbad/i
    
    excel.workbooks.each do |workbook|
      if workbook.sheets(1).range('A1').value =~ /cenbad/i && !preadvice
        preadvice = workbook.sheets(1)
      elsif (workbook.sheets(2).range('A1').value =~ /PartNumber/i rescue false)
        part_list = workbook.sheets(2)
      end
    end
 
    if preadvice.nil?
      errbox "Unable to find CENBAD Manifest\nMust have CENBAD in cell A1"
      return false
    elsif part_list.nil?
      errbox 'Unable to find Item by Account List'
      return false
    end
    
    # Determine warranty status
    self.warranty = !!( preadvice.range('A1').value =~ /doa/i )
    puts 'Warranty' if warranty
    
    # Attach to preadvice sheet
    self.preadvice = preadvice
    expected_columns = 'Item', 'Purchase Req', 'Vendor Return', 'Repair Order', 'Description 1', 'Price', 'Fault Report'
    
    raw_data = RubyExcel::Workbook.new.load preadvice.usedrange.value
    while ( raw_data.a1.nil? || raw_data.a1 =~ /cenbad/i )  && raw_data.maxrow > 2
      raw_data.row(1).delete
    end
    
    raw_data.gc! expected_columns
    
    if raw_data.maxcol < expected_columns.length
      errbox "Columns missing!\n\nExpected:\n#{ expected_columns.join("\n") }"
      return false
    end

    raw_data.rows(2) do |r|
          r.cell_h( 'Item' ).tap { |c| c.value = c.value.to_i.to_s if c.value.is_a?( Float ) }
      [ 'Purchase Req', 'Vendor Return', 'Repair Order' ].each { |h| r.cell_h(h).tap { |c| c.value = c.value.to_s.upcase.strip.force_encoding("ASCII-8BIT").gsub( /#{ "\xB6|\xFF".force_encoding("ASCII-8BIT") }/, '' ) } }
    end

    column = raw_data.ch('Repair Order')
    if column.count( data[:RP] ) == 0
      errbox "#{ data[:RP] } not found in pallet manifest."
      return false
    end
    if column.count( data[:RP] ) > 1
      errbox "#{ data[:RP] } occurs more than once."
      return false
    end

    rownum = raw_data.match( 'Repair Order', &/#{ data[:RP] }/ )
    if rownum.nil?
      errbox "#{ data[:RP] } not found in manifest."
      return false 
    else
      line = raw_data.row( rownum )
    end
    line[ 1 ] = line[ 1 ].to_i.to_s if is_float? line[ 1 ]
    
    if preadvice.cells.find( data[:RP] ).interior.color == 12_566_463
      errbox "#{ data[:RP] } has already been booked in!"
      return false 
    end
    
    unless line[ 1 ] == ( data[:Item] )
      errbox "Wrong Item!\nScanned: #{ data[:Item] }\nManifest: #{line[ 1 ]}"
      return false 
    end

    data[:PR], data[:VR] = line[ 2..3 ]
    unless data[:PR][0..1] == 'PR' && data[:VR][0..1] == 'VR'
      errbox "Invalid PR or VR reference"
      return false
    end
    
    if raw_data.ch( 'Purchase Req' ).count( data[:PR] ) != 1
      errbox "Duplicate PR reference: #{ data[:PR] }"
      return false
    end
    
    data[:Desc], data[:Cost] = line.val( 'Description 1' ).to_s, line.val( 'Price' ).to_f
    
    if data[:Desc].empty?
      errbox "Missing Description: #{ data[:PR] }"
      return false
    end
    
    if data[:Cost].zero? && !warranty
      errbox "Missing Price: #{ data[:PR] }"
      return false
    elsif !data[:Cost].zero? && warranty
      errbox "Expected warranty but price is #{ data[:Cost] }: #{ data[:PR] }"
      return false
    end
    
    data[:Fault] = warranty ? line.val( 'Fault Report' ) : nil

    part_sheet = RubyExcel::Workbook.new.load part_list.usedrange.value
    part_sheet.gc! 'PartNumber', 'Team', 'Repair Destination'
    part_sheet.rows(2) { |r| r.map! { |v| v.nil? ? nil : ( is_float?( v ) ? v.to_i.to_s : v.to_s.upcase  ) } }
    rownum =  part_sheet.match('PartNumber', &/^#{ data[:Item] }$/i )
    if rownum.nil?
      # Part doesn't exist!
      
      # Prompt for account
      data[:Account] = 
      inputbox( "Please enter the account number for #{ data[:Item] }\n#{ data[:Desc] }", 'Account', data[:Account] ).tap do |res|
        if res.nil?
          errbox "Account selection cancelled by user."
          return false
        elsif res.empty? || res.strip !~ /^PHX\d{2}$/i
          errbox "Invalid Account: #{ res }"
          return false
        end
      end.upcase.strip
      
      destinations = 'HUB MK', 'HUB NOTTINGHAM', 'DMS'
      
      # Prompt for destination
      data[:Destination] = 
      destinations[ Integer( inputbox( "Please select the Destination for #{ data[:Item] }\n#{ destinations.map.with_index { |d, i| (i+1).to_s + ' - ' + d }.join($/) }", 'Destination' ).tap do |res|
        if res.nil?
          errbox "Destination selection cancelled by user."
          return false
        elsif res.strip !~ /^[1-4]$/i
          errbox "Invalid Destination, must select 1, 2, or 3: #{ res }"
          return false
        end
      end )-1 ]
      
      # Add to Part List
      part_list.range("A#{ lastrow = part_list.usedrange.rows.count + 1 }:H#{ lastrow }").value = [ [ data[:Item], data[:Account], nil, nil, nil, data[:Desc], nil, data[:Destination] ] ]
      part_list.parent.save
      
    elsif part_sheet.row(rownum).val( 'Team' ).nil?
      errbox "Account not found for #{ data[:Item] }."
      return false
    elsif part_sheet.row(rownum).val( 'Repair Destination' ).nil?
      errbox "Destination not found for #{ data[:Item] }."
      return false
    else
      data[:Account] = part_sheet.row(rownum).val( 'Team' ).to_s
      data[:Destination] = part_sheet.row(rownum).val( 'Repair Destination' ).to_s
    end
    puts data.map {|e| e.join(': ') }.join($/)

    true
  end

  def created_warranty_item
    MechReporter.new.send_email( [], 'Warranty Item Created', ( test ? Passwords::MyName : Passwords::WarrantyCreationList ), "Item Type: #{ data[:Item] }\nDescription: #{ data[:Desc] }\nAccount: #{ data[:Account] }\nCost: #{ data[:Cost] }" )
  end
  
  def await_input
    [
      [ 'Enter the Repair Order', 'RP', data[:RP] ],
      [ 'Enter the Item Number', 'Item', data[:Item] ],
      [ 'Enter the Serial Number', 'Serial', data[:Serial] ] 
    ].each do |q|
      result = inputbox( *q)
      raise StandardError, 'User Quit' if q[1] == 'RP' && result.nil?
      return false if result.nil? || result == ''
      result.force_encoding"ASCII-8BIT"
      result = result.sub("\x9C".force_encoding("ASCII-8BIT"),'#').gsub(/[#{ "\xB6|\xFF".force_encoding("ASCII-8BIT") }@"'].+/,'').upcase.strip
      data[q[1].to_sym] = result
    end
    true
  end

  def check_page
    begin
      start( true, true, true ) if driver.nil?
    rescue
      errbox "Unable to acquire Firefox profile.\nPlease close Firefox to avoid consuming extra RDT licenses."
      return false
    end
    driver.window(:url => /pid=779/ ).use rescue login( 'Booking In' )
    true
  end
  
  def create_item_type( item: data[:Item], desc: data[:Desc], cost: data[:Cost], account: data[:Account] )
  
    # Look up the item type
    login 'Item Types'
    driver.text_field( :name => 'sp' ).set item
    driver.button( :name => 'force' ).click
    driver.text_field( :name => 'name' ).wait_until_present
    
    # Stop if it already exists
    if driver.table( :index => 2 ).rows.count > 2
      if driver.table( :index => 2 ).rows.count == 3 && driver.table( :index => 2 ).tr( index: 1 ).td( index: 1 ).text[0..2] == 'SW-'
        # Carry on
      else
        puts 'Existing Item Type found in Search!'
        return false
      end
    end
    
    puts 'Creating item: ' + item
    
    # Create it
    driver.text_field( :name => 'name' ).set item
    driver.text_field( :name => 'description' ).set desc
    driver.text_field( :name => 'price' ).set ( '%0.2f' % cost )
    driver.text_field( :name => 'name' ).parent.parent.cells.last.images.first.click
    
    # Set its attributes
    driver.checkbox(:name => 'use_serial_number').when_present.set
    driver.div(:id => 'details').table(:index => 0).trs.last.td(:index => 0).buttons[0].click
    driver.link(:title => 'Accounts').when_present.click
    driver.div(:id => 'accounts').table(:index => 0).td( :text => account ).parent.td(:index => 0).checkbox(:index => 0).set
    
    created_warranty_item if warranty
    
    true
  end
  
  def book_in_job
  
    begin # Catch Serial Number Errors
    
      begin # Catch Item Type Errors
    
        driver.div(:id => 'pldinfo').wait_while_present
        driver.text_field(:id => 'currentcustomer').set data[:Account]
        until driver.text_field(:id => 'customer_account_num').value != ''
          driver.text_field(:id => 'currentcustomer').send_keys :tab
          sleep 0.5
        end
        driver.button(:value => 'Copy To Customer Address ').click
        driver.select_list(:id => 'orderparceltype_1').select data[:Item]
        
      rescue Watir::Exception::NoValueFoundException
        
        create_item_type || fail( NameError, "Unable to create item: #{ data[:Item] }" )
        driver.execute_script('window.onbeforeunload = null') rescue nil
        check_page
        retry
        
      end
      
      Watir::Wait.until { driver.text_field( :id => 'ordertimeonsite' ).value !~ /CAL/ }
    
      driver.text_field(:id => 'orderserialnum_1').set data[:Serial]
    
    rescue Watir::Exception::ObjectDisabledException
    
      if driver.text_field(:id => 'orderserialnum_1').disabled?
        if serialize
          driver.execute_script('window.onbeforeunload = null') rescue nil
          check_page
          retry
        else
          errbox "#{ data[:Item] } is not serialised!"
          return false
        end
      end
    
    end

    driver.text_field(:id => 'ordercustomerorderref').set data[:PR]
    driver.text_field(:id => 'orderclientnum_1').set ( warranty ? "#{ data[:RP] } Warranty" : data[:RP] )
    driver.text_field(:id => 'orderlinenum_1').set data[:VR]
    driver.checkbox(:id => 'orderwclaim_1').when_present.set if warranty
    notesfield = driver.text_field(:id => 'orderparcelnotes_1')
    notesfield.set data[:Destination] + ( warranty ? "\n - Warranty" : '' )
    if data[:Fault]
      notesfield.send_keys :enter, :enter, ' *** ', data[:Fault]
    end
    
    puts driver.text_field(:id => 'orderparcelid_1').value
    driver.execute_script('window.onbeforeunload = null') rescue nil
    wait_for_download true
    driver.button(:id => 'submitbutton').click
    raise "Unexpected alert!\n#{ driver.alert.text }" if driver.alert.exists?( 2 )
    print_label wait_for_download

  end

  def tick_off_list
    cell = preadvice.usedrange.find( data[:RP] )
    cell.entirerow.interior.color = 12_566_463
    preadvice.range("Z#{cell.row}").end(-4159).offset(0,1).value = data[:Account]
  end

  def run_loop
    until !run
      puts 'Looping'
    end
    puts error.backtrace.first
    errbox error.inspect unless error.to_s == 'User Quit'
    raise error
  end

  def run
    until await_input
      puts 'Looping'
    end
    get_excel_input or return true
    check_page or return true
    book_in_job or return true
    tick_off_list
    sleep 2
    true
  rescue => self.error
    false
  end 
  
  def serialize
    # Look up the item type
    login 'Item Types'
    driver.text_field( :name => 'sp' ).set data[:Item]
    driver.button( :name => 'force' ).click
    driver.text_field( :name => 'name' ).wait_until_present
    
    # Stop if you don't find anything
    if driver.table( :index => 2 ).rows.count < 3
      puts 'Unable to find item.'
      return false
    end

    cell = driver.table( :index => 2 ).td( :text => data[:Item] )
    
    unless cell.exists?
      puts 'Unable to find item.'
      return false
    end
    
    cell.parent.td(:index => 0).button(name: 'expand').click
    
    #Set serial attribute
    driver.checkbox(:name => 'use_serial_number').when_present.set
    driver.div(:id => 'details').table(:index => 0).trs.last.td(:index => 0).buttons[0].click
    true
  end
  
  def is_float?( val )
    !!Float(val) rescue false
  end
  
end # Booking


###################################
###################################


class RGSBooking < Booking

  def await_input
    [
      [ 'Enter the CNTX Number', 'CNTX', data[:CNTX] ],
      [ 'Enter the Serial Number', 'Serial', data[:Serial] ] 
    ].each do |q|
      result = inputbox( *q)
      raise StandardError, 'User Quit' if q[1] == 'CNTX' && result.nil?
      return false if result.nil? || result == ''
      result = cntxify( result ) if q[1] == 'CNTX'
      data[q[1].to_sym] = result.upcase
    end
    true
  end

  def check_page
    start( true, true, true ) if driver.nil?
    driver.window(:url => /pid=852/ ).use rescue login( 'Booking In Triage' )
  end

  def book_in_job
    driver.text_field(:id => 'barcode').set data[:CNTX]
    driver.button( :id => 'submit' ).click
    
    if driver.alert.exists?(2)
      text = driver.alert.text
      driver.alert.close
      errbox text
      return false
    end
    
    driver.td( id: 'itemtypebox' ).when_present.flash
    if driver.td( id: 'itemtypebox' ).text == ''
      errbox data[:CNTX] + ' has no item type!'
      return false
    end
    
    driver.button( id: 'serialnumberchangebtn' ).click
    driver.text_field( id: 'newserialnumber' ).when_present.set data[:Serial]
    driver.text_field( id: 'newserialnumber' ).flash
    
    wait_for_download true
    driver.button( id: 'storechangebtn' ).click
    print_label wait_for_download
    
    true
  end

  def run_loop
    until !run
      puts 'Looping'
    end
    puts error.backtrace
    errbox error.inspect
  end

  def run
    until await_input
      puts 'Looping'
    end
    check_page
    book_in_job
    sleep 2
    true
  rescue => self.error
    false
  end 

end # RGSBooking

###################################
###################################

class BookingStripSerials<Booking

  def initialize( test = false )
    super
    
    start( true, false, true ) if driver.nil?
    driver.window(:url => /pid=244/ ).use rescue login( 'Depot Operations' )
    
    msgbox 'Set up the Item Type, Account, and Location before scanning!'
    
    serial_request
    
  end
  
  def serial_request
    loop do
    
      historycounter = driver.span( id: 'historycounter' )
      counter = historycounter.text.to_i
    
      serial = inputbox( 'Scan The Serial Number:', 'Serial' )
      serial = case serial
      when nil
        raise StandardError, 'User Quit'
      when /\A.\z/
        error 'Serial must be longer than 1 digit'
        next
      else
        serial[0..-2]
      end
      
      driver.text_field( id: 'input_field' ).send_keys serial, :enter
      
      if driver.alert.exists?(1)
        driver.alert.close
        next
      end
      
      # Print label here
      Watir::Wait.until(5) { historycounter.text.to_i > counter } rescue next
      driver.div(id: 'historyblock').divs.first.links.first.click rescue next
      
    end
  end
  
end # BookInStock

###################################
###################################

class BTEBooking<Booking

  def get_excel_input
  
    expected_columns = 'Serial Number', 'BT Purchase Order', 'BT Unity Partcode', 'Partcode Description', 'Default cost'
  
    preadvice = nil
    
    begin
      excel = WIN32OLE::connect('excel.application')
    rescue WIN32OLERuntimeError
      errbox 'Unable to find Excel'
      return false
    end
    
    excel.workbooks.each do |workbook|
      if workbook.sheets(1).range('A1').value =~ /BT Purchase Order/i
        preadvice = workbook.sheets(1)
      end
    end
 
    if preadvice.nil?
      errbox 'Unable to find BTE Manifest\nMust have "BT Purchase Order" in cell A1 of Sheet1'
      return false
    end
    
    self.preadvice = preadvice
    
    raw_data = RubyExcel::Workbook.new.load preadvice.usedrange.value
    raw_data.gc! expected_columns
    if raw_data.maxcol < expected_columns.length
      errbox "Columns missing!\n\nExpected:\n#{ expected_columns.join("\n") }"
      return false
    end
    raw_data.rows(2) do |r|
      r.map! { |v| v.nil? ? nil : ( is_float?( v ) ? v : v.to_s.upcase.strip ) }
    end
    
    # Ensure the serial is present and only occurs once
    column = raw_data.ch('Serial Number')
    case column.count( data[:Serial] )
    when 0
      errbox "#{ data[:Serial] } not found in manifest."
      return false
    when 1
      # Do nothing
    else
      errbox "#{ data[:Serial] } occurs more than once."
      return false
    end

    rownum = raw_data.match( 'Serial Number', &/#{ data[:Serial] }/ )
    if rownum.nil?
      errbox "#{ data[:Serial] } not found in manifest."
      return false 
    else
      line = raw_data.row( rownum )
    end
    
    if preadvice.cells.find( data[:Serial] ).interior.color == 12_566_463
      errbox "#{ data[:Serial] } has already been booked in!"
      return false 
    end
    
    data[:PO], data[:Item], data[:Desc], data[:Cost] = line[2..5]
    data[:Cost] = data[:Cost].to_f
    
    data.each do |k,v|
    
      if v.nil?
        errbox "#{ k } not found in manifest."
        return false
      end
    end
    
    if raw_data.ch( 'BT Purchase Order' ).count( data[:PO] ) != 1
      errbox "Duplicate PO reference: #{ data[:PO] }"
      return false
    end
    
    data[:Item].insert( 0, 'BTE-' ) unless data[:Item] =~ /\ABTE-/
    
    puts data.map {|e| e.join(': ') }.join($/)

    data[:Account] = 'BTE01'
    
    true
  end

  def await_input
    result = inputbox( 'Enter the Serial', 'Serial', data[:Serial] )
    raise StandardError, 'User Quit' if result.nil?
    return false if result.empty?
    data[:Serial] = result.upcase.strip    
    true
  end

  def book_in_job
  
    begin
    
      driver.div(:id => 'pldinfo').wait_while_present
      driver.text_field(:id => 'currentcustomer').set 'BTE01'
      until driver.text_field(:id => 'customer_account_num').value != ''
        driver.text_field(:id => 'currentcustomer').send_keys :tab
        sleep 0.5
      end
      driver.button(:value => 'Copy To Customer Address ').click
      driver.text_field(:id => 'ordercustomerorderref').set data[:PO]
      
      driver.select_list(:id => 'orderparceltype_1').select data[:Item]
      
    rescue Watir::Exception::NoValueFoundException
    
      create_item_type || fail( NameError, "Unable to create item: #{ data[:Item] }" )
      driver.execute_script('window.onbeforeunload = null') rescue nil
      check_page
      retry
      
    end
    
    Watir::Wait.until { driver.text_field( :id => 'ordertimeonsite' ).value !~ /CAL/ }
    if driver.text_field(:id => 'orderserialnum_1').disabled?
      errbox "#{ data[:Item] } is not serialised!"
      return false 
    end
    driver.text_field(:id => 'orderserialnum_1').set data[:Serial]
    
    puts driver.text_field(:id => 'orderparcelid_1').value
    driver.execute_script('window.onbeforeunload = null') rescue nil
    wait_for_download true
    
    driver.button(:id => 'submitbutton').click   
    raise "Unexpected alert!\n#{ driver.alert.text }" if river.alert.exists?( 2 )

    print_label wait_for_download
  end
  
  def tick_off_list
    cell = preadvice.usedrange.find( data[:Serial] )
    cell.entirerow.interior.color = 12_566_463
  end

end # BTEBooking