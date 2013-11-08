require 'gtk2' # GUI
require_relative '../MechReporter' # Reporter and link to RubyExcel

#
# Graphical interface for MechReporter
#

class LabelGui < Gtk::Window
  include RegistryTools

  # Build the GUI
  def initialize
  
    # Work out where list.txt lives
    @filename = get_filename
  
    # Create the main window
    window = Gtk::Window.new
    window.set_title 'RDT Label Printer'
    
    # Create a table: height, width
    table = Gtk::Table.new 3, 3, true
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
      
      # Let's do this thing!
      m = MechReporter.new
      keys = get_key.map { |k| m.cntxify( k ) }
      m.print_prompt( m.save_labels( keys, label_selected ) )
      
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
  
end

# Allows EXE builds without showing the GUI
MechReporter.new if defined?( Ocra )

# Let there be a GUI!
GUI = LabelGui.new

# A bit of devious metaprogramming to pass through messages from the reporter
class MechReporter
  def puts( str )
    GUI.gui_puts str, false
  end
end

# Go!
GUI.run