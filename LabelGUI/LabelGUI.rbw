require 'gtk2' # GUI
require_relative '../MechReporter' # Reporter and link to RubyExcel

#
# Graphical interface for MechReporter
#

class LabelGui < Gtk::Window
  include RegistryTools
  include DateTools

  # Build the GUI
  def initialize
  
    # Work out where list.txt lives
    @filename = get_filename
  
    # Create the main window
    window = Gtk::Window.new
    window.set_title 'Label Printer'
    
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
    
    # Button to close the program
    close_button = Gtk::Button.new '_Close'
    table.attach close_button, 2, 3, 0, 1
    close_button.signal_connect('clicked') { Gtk.main_quit }

    # Radio buttons to select the label type
    b1 = Gtk::RadioButton.new( 'Centrex Label' )
    b2 = Gtk::RadioButton.new( b1, 'Stock Label' )
    b3 = Gtk::RadioButton.new( b1, 'Repair Label' )

    @label_radio = b1, b2, b3
    @label_radio[ ( get_regkey_val('labeltype') || 0 ).to_i ].set_active(true)
    @label_radio.each_with_index { |b,i|
      table.attach( b, i, i+1, 1, 2 )
      b.signal_connect('toggled') { set_regkey_val( 'labeltype', i ) if b.active? }
    }
    
    # Label to update the user with feedback
    putsframe = Gtk::Frame.new
    @puts = Gtk::Label.new( 'Waiting for user input...' )
    table.attach putsframe, 0, 3, 2, 3
    putsframe.add @puts
    
    # Checkbox to instantly print
    @instaprint = Gtk::CheckButton.new 'Instant Print'
    @instaprint.active = get_regkey_val( 'instaprint' )
    table.attach @instaprint, 0, 2, 3, 4
    @instaprint.signal_connect('toggled') { set_regkey_val( 'instaprint', @instaprint.active?.to_s )  }

    # Window management stuff. No touchy!
    window.signal_connect('destroy') { Gtk.main_quit }
    window.show_all
    
  end
  
  # Display the GUI
  def run
    Gtk.main
  end
  
  def label_selected
    @label_radio.each_with_index { |b,i| return i+1 if b.active? }
  end
  
  # Let there be a printing of labels!
  def go_go_go
  
    # Catch errors so the GUI doesn't just mysteriously disappear
    begin
    
      # Make sure they have a RDT username and password
      unless get_regkey_val('username') && get_regkey_val('password')
        return false unless ask_for_details( 'RDT Details:' )
      end
      
      # Make sure there's a list.txt to use
      unless File.exist?( @filename )
        error 'Unable to find ' + @filename
        return false
      end
      
      gui_puts 'Please Wait'
      
      q = RDTQuery.new 'http://centrex.redetrack.com/redetrack/bin/report.php?report_locations_list=T&select_div_last_shown=&report_limit_to_top_locations=N&action=ordtrack&num=OR0000857779&num_raisedtrack=&status=999&pod=A&status_code=&itemtype=&location=&value_location=&tf=current&days=365&befaft=b&dd=09&mon=08&yyyy=2013&fdays=1&fbefaft=a&fdd=09&fmon=08&fyyyy=2013&ardd=09&armon=08&aryyyy=2012'
      q.set_dates( today )
      
      r = RDTQuery.new 'http://centrex.redetrack.com/redetrack/bin/report.php?report_locations_list=T&select_div_last_shown=&report_limit_to_top_locations=N&action=custordtrack&num=INC000000814606&num_raisedtrack=&status=&pod=A&status_code=&itemtype=&location=&value_location=&tf=current&days=365&befaft=b&dd=23&mon=08&yyyy=2013&fdays=1&fbefaft=a&fdd=23&fmon=08&fyyyy=2013&ardd=23&armon=08&aryyyy=2012'
      r.set_dates( today )
      
      s = RDTQuery.new 'http://centrex.redetrack.com/redetrack/bin/report.php?report_locations_list=T&select_div_last_shown=&report_limit_to_top_locations=N&action=stockstatus&num=325-333-814&num_raisedtrack=&status=&pod=A&status_code=&itemtype=&location=&value_location=&tf=current&days=365&befaft=b&dd=26&mon=10&yyyy=2013&timetype=any&fdays=1&fbefaft=a&fdd=26&fmon=10&fyyyy=2013&ardd=26&armon=10&aryyyy=2012'
      s.set_dates( today )
      
      # Let's do this thing!
      m = MechReporter.new
      
      # We're going to take our list of CNTX and OR numbers and only end up with CNTX numbers
      keys = get_key.map do |k|
        
        # If it's an OR number
        if k =~ /^OR/i
          
          # Look up the "Returns" against this Order
          q[ 'num' ] = k
          res = m.run( q )
          
          # If there's no result, map nil
          if res.maxrow == 1
            nil
            
          # If there's a result, map the OR into all the CNTX numbers.
          else
            res.ch( 'Bar Code' ).each_wh.to_a
          end
        
        elsif k =~ /^INC/i
          
          # Look up the "Returns" against this Order
          r[ 'num' ] = k
          res = m.run( r )
          
          # If there's no result, map nil
          if res.maxrow == 1
            nil

          # If there's a result, map the OR into all the CNTX numbers.
          else
            res.ch( 'Bar Code' ).each_wh.to_a
          end
        
        # If it's a CNTX or STK, leave it alone
        elsif k =~ /^CNTX|^STK/i
        
          k

        # If its a serial number
        else
          s[ 'num' ] = k
          res = m.run( s )
          if res.empty?
            nil
          else
            res.last_row.val('Bar Code')
          end
          #raise ArgumentError, 'Invalid CNTX / OR Number: ' + k
        end
        
      end.flatten.compact
      
      # Get the labels and open the print dialog
      if @instaprint.active?
        m.print_label( m.save_labels( keys, label_selected ) )
      else
        m.print_prompt( m.save_labels( keys, label_selected ) )
      end
      
      sleep 2
      
    # Catch and report errors
    rescue => e
      error "#{ e.to_s.scan(/(.{1,200}\s?.{1,200}+)/).first.join($/) }\n\n#{ e.backtrace.first }"
      gui_puts 'Error!'
    else
      gui_puts 'Task Complete.'
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
    Win32::Registry::HKEY_CURRENT_USER.open('SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders')['Personal'] + '\\labelGUI.txt'
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
  
end

# Allows EXE builds without showing the GUI
MechReporter.new if defined?( Ocra )

# Let there be a GUI!
GUI = LabelGui.new

class MechReporter
  # A bit of devious metaprogramming to pass through messages from the reporter
  def puts( str )
    GUI.gui_puts str, false
  end
end

# Go!
GUI.run
