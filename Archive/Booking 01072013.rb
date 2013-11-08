require_relative 'RDT'
require_relative 'WinBoxes'
class Booking < RDT
  include WinBoxes

  attr_accessor  :data

  def initialize( test = false, warranty = false )
    super(test)
    @test = test
    @data = {}
    @warranty = warranty
    temp = %w{username password}
    @@user_pass.map!.with_index { |val, idx| puts "Please enter RDT #{ temp[idx] }:"; exit if set_regkey_val( temp[idx], gets.chomp ) == '' } unless @@user_pass[0] && @@user_pass[1]

  end

  def get_excel_input
    preadvice, part_list = nil, nil
    excel = WIN32OLE::connect('excel.application')
    @excel = excel
    excel.workbooks.each do |workbook|
      if workbook.sheets(1).range('A1').value =~ /cenbad/i
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
    
    @preadvice = preadvice
    
    raw_data = Excel_Sheet.new preadvice.usedrange.value
    raw_data = raw_data.skip_headers
    raw_data.get_cols! ['Item', 'Purchase Req', 'Vendor Return', 'Repair Order']
    if raw_data[0].nil? || raw_data[0].length != 4
      errbox "Columns missing! Expected: Item | Purchase Req | Vendor Return | Repair Order"
      return false 
    end
    raw_data.upcase!
    column = raw_data.get_col('Repair Order')
    if column.count( @data[:RP] ) == 0
      errbox "#{@data[:RP]} not found in pallet manifest."
      return false
    end
    if column.count( @data[:RP] ) > 1
      errbox "#{@data[:RP]} occurs more than once."
      return false
    end

    line = raw_data.select { |ar| ar[3] == @data[:RP] }.flatten
    if line.empty?
      errbox "#{@data[:RP]} not found in manifest."
      return false 
    end
    line[0] = line[0].to_i.to_s if is_float? line[0]
    
    if @preadvice.cells.find(@data[:RP]).interior.color == 12_566_463
      errbox "#{@data[:RP]} has already been booked in!"
      return false 
    end
    
    unless line[0] == ( @data[:Item] )
      errbox "Wrong Item!\nScanned: #{@data[:Item]}\nManifest: #{line[0]}"
      return false 
    end

    @data[:PR], @data[:VR] = line[1].strip.force_encoding("ASCII-8BIT").gsub( /#{ "\xB6|\xFF".force_encoding("ASCII-8BIT") }/, '' ), line[2].strip.force_encoding("ASCII-8BIT").gsub( /#{ "\xB6|\xFF".force_encoding("ASCII-8BIT") }/, '' )
    unless @data[:PR][0..1] == 'PR' && @data[:VR][0..1] == 'VR'
      errbox "Invalid PR or VR reference"
      return false
    end
    part_sheet = Excel_Sheet.new(part_list.usedrange.value)
    part_sheet.get_cols!( %w(ItemNumber Team) )
    part_sheet.upcase!
    if part_sheet.get_col('PARTNUMBER').index(@data[:Item]).nil?
      errbox "#{@data[:Item]} not found in Item / Team list."
      return false
    end
    @data[:Account] = part_sheet.get_value( 'TEAM', 'PARTNUMBER', @data[:Item] )
    puts @data.map {|e| e.join(': ') }.join($/)

    true
  end

  def await_input
    [
      [ 'Enter the Repair Order', 'RP', @data[:RP] ],
      [ 'Enter the Item Number', 'Item', @data[:Item] ],
      [ 'Enter the Serial Number', 'Serial', @data[:Serial] ] 
    ].each do |q|
      result = inputbox( *q)
      raise StandardError, 'User Quit' if q[1] == 'RP' && result.nil?
      return false if result.nil? || result == ''
      result.force_encoding"ASCII-8BIT"
      result = result.sub("\x9C".force_encoding("ASCII-8BIT"),'#').gsub(/[#{ "\xB6|\xFF".force_encoding("ASCII-8BIT") }@"'].+/,'').upcase.strip
      @data[q[1].to_sym] = result
    end
    true
  end

  def check_page
    @driver ||= start( true, true, true )
    @driver.window(:url => /pid=779/ ).use rescue login( 'Booking In' )
  end

  def book_in_job
    @driver.div(:id => 'pldinfo').wait_while_present
    @driver.text_field(:id => 'currentcustomer').set @data[:Account]
    until @driver.text_field(:id => 'customer_account_num').value != ''
      @driver.text_field(:id => 'currentcustomer').send_keys :tab
      sleep 0.5
    end
    @driver.button(:value => 'Copy To Customer Address ').click
    @driver.text_field(:id => 'ordercustomerorderref').set @data[:PR]
    @driver.select_list(:id => 'orderparceltype_1').select @data[:Item]
    Watir::Wait.until { @driver.text_field( :id => 'ordertimeonsite' ).value !~ /CAL/ }
    @driver.text_field(:id => 'orderclientnum_1').set ( @warranty ? "#{ @data[:RP] } Warranty" : @data[:RP] )
    @driver.text_field(:id => 'orderserialnum_1').set @data[:Serial]
    @driver.text_field(:id => 'orderlinenum_1').set @data[:VR]
    @driver.text_field(:id => 'orderparcelnotes_1').set 'Warranty' if @warranty
    @driver.checkbox(:id => 'orderwclaim_1').set if @warranty
    
    puts @driver.text_field(:id => 'orderparcelid_1').value
    @driver.execute_script('window.onbeforeunload = null') rescue nil
    wait_for_download true
    @driver.button(:id => 'submitbutton').click
    print_label wait_for_download
  end

  def tick_off_list
    cell = @preadvice.usedrange.find( @data[:RP] )
    cell.entirerow.interior.color = 12_566_463
    @preadvice.range("Z#{cell.row}").end(-4159).offset(0,1).value = @data[:Account]
  end

  def run_loop
    until !run
      puts 'Looping'
    end
    puts @error.backtrace.first
    errbox @error.inspect
  end

  def run
    until await_input
      puts 'Looping'
    end
    return true unless get_excel_input
    check_page
    book_in_job
    tick_off_list
    sleep 2
    true
  rescue => error
    @error = error
    false
  end 
  
  def is_float?( val )
    !!Float(val) rescue false
  end
  
end
