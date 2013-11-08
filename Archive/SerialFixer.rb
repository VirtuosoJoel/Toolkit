require_relative 'RDT'
require_relative 'WinBoxes'

#
# Obsolete class
#

class SerialFixer < RDT
  include WinBoxes

  attr_accessor  :data

  def initialize( test = false )
    super(test)
    @test = test
    @data = {}
    temp = %w{username password}
    @user_pass.map!.with_index { |val, idx| puts "Please enter RDT #{ temp[idx] }:"; exit if set_regkey_val( temp[idx], gets.chomp ) == '' } unless @user_pass[0] && @user_pass[1]

  end

  def await_input
    %w(CNTX Serial Part).each do |q|
      result = inputbox( "Enter the #{q}", q )
      raise StandardError, 'User Quit' if result.nil?
      return false if result == '' && q != 'Part'
      result.force_encoding"ASCII-8BIT"
      result = result.sub("\x9C",'#').gsub(/[@"'].+/,'').upcase
      @data[q.to_sym] = result
    end
    true
  end

  def get_tool( url )
    @driver ||= start( true, true, true )
    begin
      @driver.window(:url => /#{url}/ ).use
    rescue
      name = (url.include?('report.php') ? 'Reports' : 'SerialFixer' )
      @driver.execute_script("window.open();") if name == 'Reports'
      sleep 0.5
      @driver.windows.last.use
      login( name )
    end
  end

  def change_serial
    @driver.text_field(:name => 'barcode').set @data[:CNTX]
    @driver.button(:value => 'Lookup').click
    begin
      serial_field = @driver.text_field(:name => 'serial_number')
      serial_field.set @data[:Serial] if @data[:Serial] != ''
    rescue Watir::Exception::UnknownObjectException
      errbox "Job reference not found: #{@data[:CNTX]}"
      return false
    end
    
    begin
      @driver.select_list(:name => 'item_type').select_value @data[:Part] if @data[:Part] != ''
    rescue Watir::Exception::NoValueFoundException
      errbox "Part number #{@data[:Part]} not found in list."
      return false
    end
    @driver.button(:value => 'Store Change').click
    true
  end

  def save_label
    @driver.select_list( :id => 'trackselect' ).select 'Package ID'
    @driver.text_field( :id => 'num' ).set @data[:CNTX]
    @driver.select_list(:id => 'days').select '1 year'
    @driver.button( :id => 'mainsearchbutton' ).when_present.click
    @driver.checkboxes( :name => 'update[]' ).last.set
    @driver.select_list( :name => 'print_link').select( ( @data[:CNTX].include?('STK') ? 'stock label' : 'repair label' ) )
    wait_for_download true
    @driver.button( :value => 'Print Checked' ).click
    @driver.alert.close
  end

  def run_loop
    until !run
      puts 'Looping'
    end
    puts @error.backtrace#.first
    errbox @error.inspect
  end

  def run
    until await_input
      puts 'Looping'
    end
    get_tool 'barcodedetails'
    return true unless change_serial
    get_tool 'report.php'
    save_label
    print_label wait_for_download
    sleep 2
    true
  rescue => error
    @error = error
    false
  end 
end