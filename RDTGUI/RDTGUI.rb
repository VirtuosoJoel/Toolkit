# ToDo:
# Engineer Console
# Reports
# Unallocate
# Print Labels
# Colour list in progress
# CNTXify for everything except Reports not "action=boxtrack"

require 'gtk2' # GUI
require_relative '../RDT' # RDT Controller
require_relative '../MechReporter' # Reporter and link to RubyExcel

#
# Modification to simplify ComboBox population.
#

class Gtk::ComboBox
  def <<( other )
    if other.respond_to? :each
      other.each { |item| append_text item }
    else
      append_text other
    end
  end
end

#
# Modifictions to simplify ListStore usage
#

class Gtk::ListStore
  def empty?
    get_iter('0').nil?
  end

  alias old_each each
  def each( &block )
    return to_enum :old_each unless block_given?
    old_each &block
  end
  
  def cntxify!
    each do |_, _, iter|
      iter[0] = MechReporter.cntxify( iter[0] )
    end
  end
  
end

#
# Modifications to simplify TreeIter usage
#

class Gtk::TreeIter

  def <<( *other )
    if other.length != 1
      other.each.with_index { |item, idx| self[idx] = item }
    else
      self[0], self[1] = other[0], 'black'
    end
  end

  def moose
    
  end
  
end

#
# Graphical interface for RDT
#

class RDTGUI < Gtk::Window
  include RegistryTools, DateTools
  
  # Turn string into Array
  def self.createArray( input )
    input.split(/\n+/).map{ |el| el.strip!; el.empty? ? nil : el }.compact
  end
  
  #
  # Set Constants
  #
  
  StatusList = createArray(<<-HEREDOC.gsub(/(^ +)/, ''))
    Repairing
    Testing
    On Soak Test
    ROND Request
    Component Repair
    Uplift Request Email
    BER Request Email
    On Hold - Awaiting Parts
    On Hold - WPIB
    On Hold - Re-assigning
    On Hold-Awaiting Information
    On Hold - External Repair
    Returned at Customers Request
    Beyond Economical Repair
    Repair Complete
    Allocate Only
  HEREDOC
  
  ReasonList = createArray(<<-HEREDOC.gsub(/(^ +)/, ''))
    CR001 - Fix using componant / part replacement
    CR002 - Fix using a component / part repair
    CR003 - Fix (NPR) by soft / firmware re load only
    CR004 - Fix (NPR) by minor adjustment
    CR005 - Fix (NPR) by formal calibration
    CR006 - Fix (NPR) by Inspect clean & test
    CR007 - Fix (NPR) by foreign object removal
    CR009 - Fix via external repair
    CR010 - Fix (NPR) Inspect clean & test Toner/ink
    CR011 - Fix (NPR) resoldering / rework
    CR012 - Fix by whole unit replacement
    CR013 - Fix by sub assembly / module replacement
    CR014 - No fault found
    CR015 - unrepaired quote / repair not accepted
    CR016 - Failed ( OOS) - Software fault
    CR017 - Fix multiple issues
    CR018 - Fix chargeable Repair - Missing items
  HEREDOC
  
  SLAList = createArray(<<-HEREDOC.gsub(/(^ +)/, ''))
    CR019 - Sourcing Difficult Parts
    CR020 - Extended Diagnostic Soak Test
    CR021 - Awaiting Technical Information
    CR022 - Awaiting Information from Customer
    CR023 - Awaiting Full Consignment
    CR024 - Awaiting Test Rig / Consumables etc
    CR025 - Required External Repair
    CR026 - Customer Re-arranging Priority
    CR027 - Awaiting Quotation Approval
    CR028 - Awaiting Secondary Part 
    CR029 - Returned to Manufacturer Under Warranty
  HEREDOC

  BERList = createArray(<<-HEREDOC.gsub(/(^ +)/, ''))
    CR031 - BER - Parts not available
    CR032 - BER - Damage
    CR033 - BER - Missing
    CR034 - BER - Transit Damage
  HEREDOC

  LabelList = createArray(<<-HEREDOC.gsub(/(^ +)/, ''))
    centrex label
    stock label
    repair label
  HEREDOC

  EngineersConsoleHelp = <<-HEREDOC.gsub(/(^ +)/, '')
    EngineersConsole:
    Scan the job barcodes into the Data Entry field.
    Click the "Update List" button to use the contents of Data Entry.
    Select a Status to apply to each job.
    If Repair Complete, select the relevant Reason and SLA Codes.
    If Beyond Economical Repair, select the relevant BER Code.
    Enter the Job Notes into the "Notes" field.
    Click "Go".
  HEREDOC
  
  ReportsHelp = <<-HEREDOC.gsub(/(^ +)/, '')
    Reports:
    Scan the job barcodes into the Data Entry field.
    Click the "Update List" button to use the contents of Data Entry.
    Run a report on RDT for one reference.
    Paste the Report URI into the URI field.
    Click "Go".
    A Report should appear in Excel.
  HEREDOC
  
  UnallocateHelp = <<-HEREDOC.gsub(/(^ +)/, '')
    Unallocate:
    Scan the job barcodes into the Data Entry field.
    Click the "Update List" button to use the contents of Data Entry.
    Click "Go".
  HEREDOC
  
  LabelsHelp = <<-HEREDOC.gsub(/(^ +)/, '')
    Print Labels:
    Scan the job barcodes into the Data Entry field.
    Click the "Update List" button to use the contents of Data Entry.
    Click "Go".
    Acrobat should open and ask which printer to print to.
  HEREDOC
  
  Column2 = []
  
  # Build the GUI
  def initialize
  
    # Create a logfile
    logfilepath = ENV['APPDATA'] + '\RDTGUI'
    Dir.mkdir( logfilepath ) unless File.exists?(logfilepath)
    $stdout.reopen( logfilepath + '\Rubylog.txt', 'w')
    $stdout.sync = true
    $stderr.reopen($stdout)
  
    # Find the current list file.
    @filename = get_regkey_val( 'rdtguilist' ) || get_filename
  
    # Create the main window
    @window = Gtk::Window.new
    @window.title = 'RDTGUI'
    
    # Create a table: height, width, homogenous
    table = Gtk::Table.new 18, 5, false
    @window.add table
    table.row_spacings = 5
    table.column_spacings = 5
    
    #
    # Top Menus
    #
    
    # Menu bar
    mb = Gtk::MenuBar.new
    table.attach mb, 0, 5, 0, 1
    
    # Menu 1 - Options
    mb.append( Gtk::MenuItem.new( 'Options' ).tap do |menutitle|
      Gtk::Menu.new.tap do |submenu|  
        menutitle.set_submenu submenu
      
        submenu.append Gtk::MenuItem.new( 'Choose List File' ).tap { |item| item.signal_connect( 'activate' ) { set_list_file } }
        submenu.append Gtk::MenuItem.new( 'Open Firefox' ).tap { |item| item.signal_connect( 'activate' ) { open_firefox } }
        submenu.append Gtk::MenuItem.new( 'Set User' ).tap { |item| item.signal_connect( 'activate' ) { ask_for_details } }
        submenu.append Gtk::MenuItem.new( 'View Logfile' ).tap { |item| item.signal_connect( 'activate' ) { Thread.new { `notepad "#{ logfilepath + '\Rubylog.txt' }"` } } }
        
      end
    end )
    
    # Menu 2 - Engineer Presets
    mb.append( Gtk::MenuItem.new( 'Engineer Presets' ).tap do |menutitle|
      Gtk::Menu.new.tap do |submenu|  
        menutitle.set_submenu submenu
      
        submenu.append Gtk::MenuItem.new( 'Testy1' ).tap { |item| item.signal_connect( 'activate' ) { set_colour( 0, %w(pink red green black blue purple).sample ) } }
        submenu.append Gtk::MenuItem.new( 'Testy2' ).tap { |item| item.signal_connect( 'activate' ) { set_colour( 1, %w(pink red green black blue purple).sample ) } }
        submenu.append Gtk::MenuItem.new( 'Testy3' ).tap { |item| item.signal_connect( 'activate' ) { set_colour( 2, %w(pink red green black blue purple).sample ) } }
        
      end
    end )

    # Menu 3 - Report Presets
    mb.append( Gtk::MenuItem.new( 'Report Presets' ).tap do |menutitle|
      Gtk::Menu.new.tap do |submenu|  
        menutitle.set_submenu submenu
      
        submenu.append Gtk::MenuItem.new( 'Barcode - Full' ).tap { |item| item.signal_connect( 'activate' ) { @uri_field.text = RDTQuery.new( RDTQuery::BCF ).set_dates( today ).to_s } }
        submenu.append Gtk::MenuItem.new( 'Barcode - Latest' ).tap { |item| item.signal_connect( 'activate' ) { @uri_field.text = RDTQuery.new( RDTQuery::BCL ).set_dates( today ).to_s } }
        submenu.append Gtk::MenuItem.new( 'Workshop - Full' ).tap { |item| item.signal_connect( 'activate' ) { @uri_field.text = RDTQuery.new( RDTQuery::WSF ).set_dates( today ).to_s } }
        submenu.append Gtk::MenuItem.new( 'Workshop - Latest' ).tap { |item| item.signal_connect( 'activate' ) { @uri_field.text = RDTQuery.new( RDTQuery::WSL ).set_dates( today ).to_s } }
        
      end
    end )
    
    # Menu 4 - Help
    mb.append( Gtk::MenuItem.new( 'Help' ).tap do |menutitle|
      Gtk::Menu.new.tap do |submenu|  
        menutitle.set_submenu submenu
      
        submenu.append Gtk::MenuItem.new( 'EngineersConsole' ).tap { |item| item.signal_connect( 'activate' ) { infobox EngineersConsoleHelp } }
        submenu.append Gtk::MenuItem.new( 'Reports' ).tap { |item| item.signal_connect( 'activate' ) { infobox ReportsHelp } }
        submenu.append Gtk::MenuItem.new( 'Unallocate' ).tap { |item| item.signal_connect( 'activate' ) { infobox UnallocateHelp } }
        submenu.append Gtk::MenuItem.new( 'Print Labels' ).tap { |item| item.signal_connect( 'activate' ) { infobox LabelsHelp } }
        
      end
    end )

    #
    # Column 1
    #
    
    # Row 2 - Label - Task
    Gtk::Frame.new.tap do |h|
      table.attach h, 0, 1, 1, 2
      h.add Gtk::Label.new( 'Task' )
    end

    # Row 3 - RadioButton - EngineersConsole
    @radiobuttons = []
    @radiobuttons << Gtk::RadioButton.new( 'EngineersConsole' )
    table.attach @radiobuttons.last, 0, 1, 2, 3
    
    # Define a special method so we can see the text of the selected item
    @radiobuttons.define_singleton_method( :selected ) { find( &:active? ).label }
    
    # Row 4 - Label - Notes
    Gtk::Frame.new.tap do |h|
      table.attach h, 0, 1, 3, 4
      h.add Gtk::Label.new( 'Notes' )
    end
    
    # Rows 5-6 - TextView - Notes
    Gtk::Frame.new.tap do |h|
      h.add @notes_field = Gtk::TextView.new
      table.attach h, 0, 1, 4, 6
      @notes_field.buffer.text = get_regkey_val( 'Notes' )
    end
    
    # Row 9 - RadioButton - Reports
    @radiobuttons << Gtk::RadioButton.new( @radiobuttons[0], 'Reports' )
    table.attach @radiobuttons.last, 0, 1, 8, 9
    
    # Row 11 - RadioButton - Unallocate
    @radiobuttons << Gtk::RadioButton.new( @radiobuttons[0], 'Unallocate' )
    table.attach @radiobuttons.last, 0, 1, 10, 11

    # Row 12 - RadioButton - Print Labels
    @radiobuttons << Gtk::RadioButton.new( @radiobuttons[0], 'Print Labels' )
    table.attach @radiobuttons.last, 0, 1, 12, 13
    
    # RadioButtons Group
    @radiobuttons.each { |b| b.signal_connect( 'clicked' ) { radio_changed @radiobuttons.selected } }
    
    # Row 14 - Checkbox - Test Server
    @test = Gtk::CheckButton.new 'Test Server'
    @test.active = get_regkey_val( 'Test' ) == 'true'
    table.attach @test, 0, 1, 13, 14
    @test.signal_connect('toggled') { set_regkey_val( 'Test', @test.active?.to_s ) }
    
    # Row 15 - Checkbox - Supervisor Console
    @supervisor = Gtk::CheckButton.new 'Supervisor Console'
    @supervisor.active = get_regkey_val( 'Supervisor' ) == 'true'
    table.attach @supervisor, 0, 1, 14, 15
    @supervisor.signal_connect('toggled') { set_regkey_val( 'Supervisor', @supervisor.active?.to_s ) }
    
    # Row 16 - Checkbox - Stop On Error
    @errorstop = Gtk::CheckButton.new 'Stop On Error'
    @errorstop.active = get_regkey_val( 'errorstop' ) == 'true'
    table.attach @errorstop, 0, 1, 15, 16
    @errorstop.signal_connect('toggled') { set_regkey_val( 'errorstop', @errorstop.active?.to_s ) }
    
    # Row 17 - ProgressBar
    @progress = Gtk::ProgressBar.new
    table.attach @progress, 0, 5, 16, 17
    
    # Row 18 - Label - Status
    Gtk::Frame.new.tap do |h|
      table.attach h, 0, 5, 17, 18
      h.add ( @status = Gtk::Label.new( 'Awaiting Input' ) )
    end
    
    #
    # Column 2
    #
    
    Column2 << @notes_field
    
    # Row 2 - Label - Details
    Gtk::Frame.new.tap do |h|
      table.attach h, 1, 2, 1, 2
      h.add Gtk::Label.new( 'Details' )
    end

    # Row 3 - ComboBox - Status
    @StatusList = Gtk::ComboBox.new
    @StatusList << StatusList
    @StatusList.signal_connect( 'changed' ) { combobox_changed }
    @StatusList.active = StatusList.index( get_regkey_val( 'statuslist' ) ) || -1
    table.attach @StatusList, 1, 2, 2, 3
    Column2 << @StatusList
    
    # Row 4 - ComboBox - Reason
    @ReasonList = Gtk::ComboBox.new
    @ReasonList << ReasonList
    @ReasonList.active = ReasonList.index( get_regkey_val( 'reasonlist' ) ) || -1
    table.attach @ReasonList, 1, 2, 3, 4
    Column2 << @ReasonList
    
    # Row 5 - ComboBox - SLA
    @SLAList = Gtk::ComboBox.new
    @SLAList << SLAList
    @SLAList.active = SLAList.index( get_regkey_val( 'slalist' ) ) || -1
    table.attach @SLAList, 1, 2, 4, 5
    Column2 << @SLAList

    # Row 6 - ComboBox - BER
    @BERList = Gtk::ComboBox.new
    @BERList << BERList
    @BERList.active = BERList.index( get_regkey_val( 'berlist' ) ) || -1
    table.attach @BERList, 1, 2, 5, 6
    Column2 << @BERList
    
    # Row 8 - Label - Report URI
    Gtk::Frame.new.tap do |h|
      table.attach h, 1, 2, 7, 8
      h.add Gtk::Label.new( 'Report URI' )
    end
    
    # Row 9 - Entry - URI
    @uri_field = Gtk::Entry.new
    table.attach @uri_field, 1, 2, 8, 9
    @uri_field.text = get_regkey_val( 'defaulturi' )
    Column2 << @uri_field
    
    # Row 10 - Checkbox - Use Today's Date
    @todaysdate = Gtk::CheckButton.new 'Use Today\'s Date'
    @todaysdate.active = get_regkey_val( 'todaysdate' ) == 'true'
    table.attach @todaysdate, 1, 2, 9, 10
    @todaysdate.signal_connect('toggled') { set_regkey_val( 'todaysdate', @todaysdate.active?.to_s ) }
    Column2 << @todaysdate

    # Row 12 - Label - Label Type
    Gtk::Frame.new.tap do |h|
      table.attach h, 1, 2, 11, 12
      h.add Gtk::Label.new( 'Label Type' )
    end

    # Row 13 - ComboBox - Label Type
    @LabelList = Gtk::ComboBox.new
    @LabelList << LabelList
    @LabelList.active = LabelList.index( get_regkey_val( 'labellist' ) ) || -1
    table.attach @LabelList, 1, 2, 12, 13
    Column2 << @LabelList
    
    #
    # Column 3
    #
    
    # Row 2 - Label - Data Entry
    Gtk::Frame.new.tap do |h|
      table.attach h, 2, 3, 1, 2
      h.add Gtk::Label.new( 'Data Entry' )
    end
    
    # Rows 3-16 - TextView - Data Entry
    @data_entry = Gtk::TextView.new
    @data_entry.set_size_request 130, -1
    table.attach @data_entry, 2, 3, 2, 16
    
    #
    # Column 4
    #

    # Row 2 - Label - Current List
    Gtk::Frame.new.tap do |h|
      table.attach h, 3, 4, 1, 2
      h.add Gtk::Label.new( 'Current List' )
    end
    
    # Rows 3-16 - TreeView - Current List
    @current_list = Gtk::ListStore.new String, String
    @list_view = Gtk::TreeView.new @current_list
    @list_view.selection.mode = Gtk::SELECTION_NONE
    @renderer = Gtk::CellRendererText.new
    col = Gtk::TreeViewColumn.new 'Item', @renderer, :text => 0
    col.add_attribute @renderer, 'foreground', 1
    col.expand = true
    @list_view.set_size_request 150, -1
    @list_view.append_column col
    scroll = Gtk::ScrolledWindow.new
    scroll.set_policy(Gtk::POLICY_AUTOMATIC, Gtk::POLICY_AUTOMATIC)
    scroll.add @list_view
    table.attach scroll, 3, 4, 2, 16, Gtk::FILL, Gtk::FILL
    
    # Test TreeView
    #set_list (123456..123489).map { |n| 'CNTX' + '%010d' % n }

    #
    # Column 5
    #

    # Row 2 - Label - Buttons
    Gtk::Frame.new.tap do |h|
      table.attach h, 4, 5, 1, 2
      h.add Gtk::Label.new( 'Commands' )
    end
    
    # Row 3 - Button - Go
    go_button = Gtk::Button.new '_Go'
    table.attach go_button, 4, 5, 2, 3
    go_button.signal_connect('clicked') { go_go_go }

    # Row 4 - Button - Update List
    go_button = Gtk::Button.new '_Update List'
    table.attach go_button, 4, 5, 3, 4
    go_button.signal_connect('clicked') { update_list }

    # Row 5 - Button - Import List
    go_button = Gtk::Button.new '_Import List'
    table.attach go_button, 4, 5, 4, 5
    go_button.signal_connect('clicked') { import_list }

    # Row 6 - Button - Edit List
    go_button = Gtk::Button.new '_Edit List'
    table.attach go_button, 4, 5, 5, 6
    go_button.signal_connect('clicked') { edit_list }
    
    # Row 7 - Button - Close
    close_button = Gtk::Button.new '_Close'
    table.attach close_button, 4, 5, 6, 7
    close_button.signal_connect('clicked') { Gtk.main_quit }
    
    #
    # Initial Data Load
    #
    
    import_list
    add_name_to_title
    radio_changed @radiobuttons.selected
    
    #
    # Window management stuff. No touchy!
    #
    
    @window.signal_connect('destroy') { Gtk.main_quit }
    @window.show_all
    
  end
  
  # Display the GUI
  def run
    Gtk.main
  end
  
  def add_name_to_title
    get_regkey_val( 'username' ).tap { |n| @window.title = 'RDTGUI - ' + n if n }
  end
  
  # Write a list as data and into the chosen list file
  def set_list( list )
    @current_list.clear
    list.sort.uniq.each { |el| @current_list.append.<< el }
    File.write( @filename, list.join($/) )
  end
  
  def set_colour( index, colour='black' )
    @current_list.set_value( @current_list.get_iter( index.to_s ), 1, colour )
  end
  
  def set_list_file
    dialog = Gtk::FileChooserDialog.new( 'Select a list', nil, Gtk::FileChooser::ACTION_OPEN, nil, [Gtk::Stock::CANCEL, Gtk::Dialog::RESPONSE_CANCEL], [Gtk::Stock::OPEN, Gtk::Dialog::RESPONSE_ACCEPT] )
    if dialog.run == Gtk::Dialog::RESPONSE_ACCEPT
      @filename = dialog.filename.force_encoding( 'UTF-8' )
      set_regkey_val 'rdtguilist', @filename
      puts "filename = #@filename"
    end
    dialog.destroy
  end
  
  def set_progress( position, text='' )
    @progress.fraction = position
    @progress.text = text
  end
  
  def gui_puts( str, internal=true )
    @msg = str if internal
    str = "#@msg - #{ str }" unless internal
    puts str
    @status.text = str
    Gtk.main_iteration while Gtk.events_pending?
  end
  
  # Let there be a printing of labels!
  def go_go_go
  
    # Catch errors so the GUI doesn't just mysteriously disappear
    begin
    
      # Make sure they have a RDT username and password
      unless get_regkey_val('username') && get_regkey_val('password')
        return false unless ask_for_details( 'RDT Details:' )
      end
      
      # Make sure there's a list to use
      error( 'Empty List!' ) if @current_list.empty?
      
      # CNTXify the list (unless it's a report which doesn't take CNTX numbers)
      @current_list.cntxify! unless @radiobuttons.selected == 'Reports' && @urifield !~ /action=boxtrack/

    # Catch and report errors
    rescue => e
      gui_puts 'Error: ' + e.to_s
      error "#{ e.to_s.scan(/(.{1,200}\s?.{1,200}+)/).first.join($/) }\n\n#{ e.backtrace.first }"
    else
      gui_puts 'Task Complete.'
    end
    
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
  
  # Extract the contents of the file into an array
  def get_key
    RDTGUI.createArray File.read( @filename )
  end
  
  # Report Error To User
  def error( msg )
    md = Gtk::MessageDialog.new( nil, Gtk::Dialog::DESTROY_WITH_PARENT, Gtk::MessageDialog::WARNING, Gtk::MessageDialog::BUTTONS_CLOSE, msg.to_s )
    md.signal_connect('response') { md.destroy }
    md.run
  end
  
  # Report Information To User
  def infobox( msg )
    md = Gtk::MessageDialog.new( nil, Gtk::Dialog::DESTROY_WITH_PARENT, Gtk::MessageDialog::INFO, Gtk::MessageDialog::BUTTONS_CLOSE, msg.to_s )
    md.signal_connect('response') { md.destroy }
    md.run
  end
  
  # Set options active/inactive
  def radio_changed( label )
  
    Column2.each { |el| el.sensitive = false }
    tru = -> x { x.sensitive = true }

    case label
    when @radiobuttons[0].label
      tru[ @StatusList ]
      tru[ @notes_field ]
      combobox_changed
    when @radiobuttons[1].label
      tru[ @uri_field ]
      tru[ @todaysdate ]
    when @radiobuttons[3].label
      tru[ @LabelList ]
    end
  end
  
  # Set ComboBoxes active/inactive
  def combobox_changed
    @ReasonList.sensitive = false
    @SLAList.sensitive = false
    @BERList.sensitive = false
  
    case @StatusList.active_text
    when 'Repair Complete'
      @ReasonList.sensitive = true
      @SLAList.sensitive = true
    when 'Beyond Economical Repair'
      @BERList.sensitive = true
    end
  end
  
  # Set the current list as the contents of a file
  def import_list
    set_list get_key
  end
  
  # Enter the current list into the Data Entry field
  def edit_list
    @data_entry.buffer.text = @current_list.to_enum.map { |_,_,a| a[0] }.join($/)
  end
  
  # Set the current list to the contents of the Data Entry field
  def update_list
    set_list RDTGUI.createArray( @data_entry.buffer.text ) unless @data_entry.buffer.text.sub(/\s+/,'').empty?
  end
  
  # Update RDT Username and Password
  def ask_for_details
  
    # Create a standard dialog
    md = Gtk::Dialog.new( 'RDT Details:', nil, Gtk::Dialog::DESTROY_WITH_PARENT, [ Gtk::Stock::OK, Gtk::Dialog::RESPONSE_ACCEPT ], [Gtk::Stock::CANCEL, Gtk::Dialog::RESPONSE_REJECT] )

    # Ask for the username
    md.vbox.add Gtk::Label.new( 'Username: ' )
    username = Gtk::Entry.new
    md.vbox.add username
    
    # Ask for the password
    md.vbox.add Gtk::Label.new( 'Password: ' )
    password = Gtk::Entry.new.tap { |p| p.visibility = false; p.caps_lock_warning = true; md.vbox.add p }
    
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
    
    add_name_to_title if ret
    
    # Return true if the user hit ok, false if they hit cancel
    ret
  end
  
  def open_firefox
    gui_puts 'Opening Firefox Without Profile-Lock...'
    RDT.new.start( true, false, true )
    gui_puts 'Opened Firefox.'
  rescue
    gui_puts 'Failed to open Firefox.'
    error 'Unable to use Firefox profile. Please close Firefox.'
  end

  # Update the label we're using to communicate with the user
  def gui_puts( str, internal=true )
    str = str.to_s
    @msg = str if internal
    str = @msg + ' - ' + str unless internal
    puts str
    @status.set_text str
    Gtk.main_iteration while Gtk.events_pending?
  end
  
end

# Allows EXE builds without showing the GUI
if defined?( Ocra )
  # Let OCRA pick up the win32ole extension
  RubyExcel::Workbook.new.documents_path
  # And the cookies from Mechanize
  Mechanize.new.cookies
  # And the encryption protocol
  Crypt::Blowfish.new('1').encrypt_string('Moose')
  # And Webdriver
  Watir::Browser.new.close
  exit
end

# Let there be a GUI!
GUI = RDTGUI.new

# A bit of devious metaprogramming to pass through messages from the reporter
class MechReporter
  def puts( str )
    GUI.gui_puts str, false
  end
end

# Go!
GUI.run