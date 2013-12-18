require 'gtk2' # GUI
require_relative '../MechReporter' # Reporter and link to RubyExcel

#
# Graphical interface for MechReporter
#

class ReporterGUI < Gtk::Window
  include RegistryTools
  include DateTools

  # Workshop Reports URI keys
  WorkshopReportKeys = {
    'Item Type' => 'item_type',
    'Barcode' => 'barcode',
    'Serial Number' => 'serial',
    'Customer Order Ref' => 'custordref',
    'Account Number' => 'account',
    'Client Number' => 'clientnum'
  }

  # Build the GUI
  def initialize
  
    # Work out where list.txt lives
    @filename = get_filename
  
    # Get the last report run
    @default_uri = get_regkey_val('defaulturi') || 'http://centrex.redetrack.com/redetrack/bin/report.php?report_locations_list=T&select_div_last_shown=&report_limit_to_top_locations=N&action=boxtrack&num_raisedtrack=&status=&pod=A&itemtype=&tf=current&days=365&befaft=b&dd=07&mon=05&yyyy=2013&timetype=any&fdays=1&fbefaft=a'
  
    # Create the report query
    @current_query = RDTQuery.new
  
    # Create the main window
    window = Gtk::Window.new
    window.set_title 'RDT Bulk Reporter'
    
    # Create a table: height, width
    table = Gtk::Table.new 4, 3, true
    window.add table

    # Button to start the report
    go_button = Gtk::Button.new '_Go'
    table.attach go_button,0, 1, 0, 1
    go_button.signal_connect('clicked') { go_go_go }

    # Button to edit list.txt
    edit_button = Gtk::Button.new 'Edit _List'
    table.attach edit_button,1, 2, 0, 1
    edit_button.signal_connect('clicked') { open_list }
    
    # Text entry field for the URI
    @uri_input = Gtk::Entry.new
    table.attach @uri_input, 1, 3, 1, 2
    @uri_input.text = @default_uri
    @uri_input.select_region( 0, -1 )
    @uri_input.grab_focus

    # Label for the URI field
    uri_label = Gtk::Label.new('_Enter Report URI:', true)
    uri_label.mnemonic_widget = @uri_input
    table.attach uri_label, 0, 1, 1, 2
    
    # Button to close the program
    close_button = Gtk::Button.new '_Close'
    table.attach close_button, 2, 3, 0, 1
    close_button.signal_connect('clicked') { Gtk.main_quit }

    # Label to update the user with feedback
    putsframe = Gtk::Frame.new
    @puts = Gtk::Label.new( 'Waiting for user input...' )
    table.attach putsframe, 0, 3, 2, 3
    putsframe.add @puts
    
    # Set up a menu of options
    mb = Gtk::MenuBar.new
    submenu = Gtk::Menu.new
    menutitle = Gtk::MenuItem.new 'Workshop Report Type'
    menutitle.set_submenu submenu
  
    @menu_items = [] 
    @menu_items << Gtk::RadioMenuItem.new(@menu_items, 'Item Type')
    @menu_items << Gtk::RadioMenuItem.new(@menu_items, 'Barcode')
    @menu_items << Gtk::RadioMenuItem.new(@menu_items, 'Serial Number')
    @menu_items << Gtk::RadioMenuItem.new(@menu_items, 'Customer Order Ref')
    @menu_items << Gtk::RadioMenuItem.new(@menu_items, 'Account Number')
    @menu_items << Gtk::RadioMenuItem.new(@menu_items, 'Client Number')
    
    @menu_items.each do |item|
      submenu.append item
    end
    
    # Set the default menu item
    @menu_items[1].active = true
    
    mb.append menutitle
    table.attach mb, 0, 1, 3, 4
    
    # Set up a menu of Preset Reports
    submenu2 = Gtk::Menu.new
    menutitle = Gtk::MenuItem.new 'Preset Reports'
    menutitle.set_submenu submenu2
  
    presets = [] 
    submenu2.append Gtk::MenuItem.new('CNTX Report').tap { |item| item.signal_connect("activate") { @uri_input.text = RDTQuery.new('http://centrex.redetrack.com/redetrack/bin/report.php?report_locations_list=T&select_div_last_shown=&report_limit_to_top_locations=N&action=boxtrack&num=CNTX0001373311&num_raisedtrack=&status=&pod=A&status_code=&itemtype=&location=&value_location=&tf=current&days=365&befaft=b&dd=16&mon=12&yyyy=2013&timetype=any&fdays=1&fbefaft=a&fdd=16&fmon=12&fyyyy=2013&ardd=16&armon=12&aryyyy=2012').set_dates(today).to_s } }
    submenu2.append Gtk::MenuItem.new('PO Report').tap { |item| item.signal_connect("activate") { @uri_input.text = RDTQuery.new('http://centrex.redetrack.com/redetrack/bin/report.php?report_locations_list=T&select_div_last_shown=&report_limit_to_top_locations=N&action=custordtrack&num=PR063183&num_raisedtrack=&status=&pod=A&status_code=&itemtype=&location=&value_location=&tf=current&days=365&befaft=b&dd=16&mon=12&yyyy=2013&fdays=1&fbefaft=a&fdd=16&fmon=12&fyyyy=2013&ardd=16&armon=12&aryyyy=2012').set_dates(today).to_s } }
    submenu2.append Gtk::MenuItem.new('Workshop Job Report').tap { |item| item.signal_connect("activate") { @uri_input.text = RDTQuery.new('http://centrex.redetrack.com/redetrack/bin/centrexticketreport.php?reptype=job&item_type=&barcode=CNTX0001380912&engineer=&serial=&custordref=&includeuaj=Yes&groupbyjob=Yes&berhandle=Y&account=&clientnum=&status=-1&statuscurrent=N&days=365&nolimit=0&range=daterange&befaft=b&dd=16&mon=12&yyyy=2013&depot=Centrex+Computing+Services&action=ticketreport&go=Search').set_dates(today).to_s } }
    submenu2.append Gtk::MenuItem.new('Workshop History Report').tap { |item| item.signal_connect("activate") { @uri_input.text = RDTQuery.new('http://centrex.redetrack.com/redetrack/bin/centrexticketreport.php?reptype=hist&item_type=&barcode=CNTX0001380912&engineer=&serial=&custordref=&groupbyjob=Yes&berhandle=D&account=&clientnum=&status=-1&statuscurrent=N&days=365&nolimit=0&range=daterange&befaft=b&dd=16&mon=12&yyyy=2013&depot=Centrex+Computing+Services&action=ticketreport&go=Search').set_dates(today).to_s } }
    
    mb.append menutitle
    
    # Window management stuff. No touchy!
    window.signal_connect('destroy') { Gtk.main_quit }
    window.show_all
    
  end
  
  # Display the GUI
  def run
    puts 'starting'
    Gtk.main
  end
  
  # Run the report
  def go_go_go
  
    # Catch errors so the GUI doesn't just mysteriously disappear
    begin
    
      # Make sure they have a RDT username and password
      unless get_regkey_val('username') && get_regkey_val('password')
        return false unless ask_for_details( 'RDT Details:' )
      end
      
      # Make sure there's a base URI to use
      if @uri_input.text.empty?
        error 'You must give an example URI!'
        return false
      end
      
      # Make sure there's a list.txt to use
      unless File.exist?( @filename )
        error 'Unable to find ' + @filename
        return false
      end
      
      # Write this URI into the registry as the default one
      set_regkey_val('defaulturi', @uri_input.text)

      # Notify the user
      gui_puts 'Running Report...'
      
      # Run the report and output the results to Excel
      each_query.to_excel
      
    # Catch and report errors
    rescue => e
      error "#{ e.to_s.scan(/(.{1,200}\s?.{1,200}+)/).first.join($/) }\n\n#{ e.backtrace.first }"
    end
    
    # Quit
    #Gtk.main_quit
  end

  # Open list.txt in Notepad
  def open_list
    File.exist?( @filename ) || File.write( @filename, '' )
    Thread.new { system "notepad #{ @filename }" }
  end
  
  # Filename for ( Documents / My Documents ) list.txt
  def get_filename
    Win32::Registry::HKEY_CURRENT_USER.open('SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders')['Personal'] + '\\list.txt'
  end
  
  # Extract the contents of list.txt into an array
  def get_key
    File.read(@filename).split(/\s/).map{ |v| v.strip!; v.empty? ? nil : v }.compact
  end

  # Report an error
  def error( msg )
    md = Gtk::MessageDialog.new( nil, Gtk::Dialog::DESTROY_WITH_PARENT, Gtk::MessageDialog::WARNING, Gtk::MessageDialog::BUTTONS_CLOSE, msg )
    md.signal_connect('response') { md.destroy }
    md.run
  end
  
  # Get input
  def ask_for_details( msg )
  
    # Create a standard dialog
    md = Gtk::Dialog.new( msg, nil, Gtk::Dialog::DESTROY_WITH_PARENT, [ Gtk::Stock::OK, Gtk::Dialog::RESPONSE_ACCEPT ], [Gtk::Stock::CANCEL, Gtk::Dialog::RESPONSE_REJECT] )

    # Ask for the username
    userlabel = Gtk::Label.new 'Username: '
    md.vbox.add userlabel
    username = Gtk::Entry.new
    md.vbox.add username
    
    # Ask for the password
    passlabel = Gtk::Label.new 'Password: '
    md.vbox.add passlabel
    password = Gtk::Entry.new
    password.visibility = false
    password.caps_lock_warning = true
    md.vbox.add password
    
    # Display the dialog
    md.show_all
    
    # Handle the response
    ret = false
    md.run do |response|
      if response == Gtk::Dialog::RESPONSE_ACCEPT
        set_regkey_val 'username', username.text
        set_regkey_val 'password', password.text
        ret = true
      else
      end
      md.destroy
    end
    
    # Return true if the user hit ok, false if they hit cancel
    ret
  end

  # Update the label we're using to communicate with the user
  def gui_puts( str, internal=true )
    @msg = str if internal
    str = @msg + ' - ' + str unless internal
    puts str
    @puts.set_text str
    Gtk.main_iteration while Gtk.events_pending?
  end

  # Runs the report with the current options and returns a RubyExcel::Sheet
  # Breaks long queries into smaller chunks.
  def each_query
  
    # Set the upper character limit for the URI
    maxlen = 6000
    
    # Set up the current query
    @current_query.query = @uri_input.text
    
    # Wipe existing numbers from the URI
    @current_query[ query_type ] = ''
    
    # Determine the maximum references we can report with
    allowed_extra = maxlen - @current_query.to_s.length
    
    # Get the numbers we need to report on
    array = get_key
    
    # Create an empty sheet to populate with data
    res = RubyExcel::Workbook.new.add
    
    # Create the reporter
    @mechwarrior ||= MechReporter.new
    
    # Keep going until we've gathered all the requisite data
    until array.empty?
    
      # Wipe out the query key for each loop
      stringy = ''
    
      # Report progress back to the impatient user
      gui_puts 'Remaining: ' + array.length.to_s
      
      # Keep building the string until we hit the limit
      until array.empty? || stringy.bytesize + array.last.bytesize >= allowed_extra
    
        # Take from array, add to string
        stringy << array.pop
        stringy << '|'
      
      end
      
      # Build the finished URI
      @current_query[ query_type ] = stringy.chomp('|')
      
      # Run the query and return the data
      res << @mechwarrior.run(@current_query)
      
    end # each_query
    
    # Clean up the data, and kill off leading equals signs to prevent Excel throwing a fit
    res.rows { |r| r.map! { |v| v.to_s.gsub(/\s/, ' ').sub(/^=/,"'=") } }
    
    # Report success ( although if we close the GUI after the report the user will never see this )
    gui_puts 'Report Complete'
    
    # Return the Sheet
    res
  end
  
  # Return the right string for the search key, based on the report type
  def query_type
    search_string = @current_query.to_s
    if search_string =~ /centrexticketreport\.php/
      #return 'serial' if search_string =~ /reptype=serial/
      
      res = @menu_items.find(&:active?).label
      res.nil? ? 'barcode' : WorkshopReportKeys[ res ]
      
    else
      'num'
    end
  end
  
end

# Allows EXE builds without showing the GUI
MechReporter.new if defined?( Ocra )

# Let there be a GUI!
GUI = ReporterGUI.new

class MechReporter
  # A bit of devious metaprogramming to pass through messages from the reporter
  def puts( str )
    GUI.gui_puts str, false
  end
end

# Go!
GUI.run