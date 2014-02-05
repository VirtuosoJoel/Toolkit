require 'gtk2' # GUI
require_relative '../MechReporter' # Reporter and link to RubyExcel

#
# Scan CNTX, get Label.
#

class LabelPrinter < Gtk::Window
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
    table = Gtk::Table.new 1, 1, true
    window.add table

    # CNTX Input
    @input = Gtk::Entry.new 
    table.attach @input, 0, 1, 0, 1
    @input.signal_connect('activate') { go_go_go; @input.select_region( 0, -1 ) }

    # Window management stuff. No touchy!
    window.signal_connect('destroy') { Gtk.main_quit }
    window.show_all
    
  end
  
  # Display the GUI
  def run
    Gtk.main
  end
  
  # Let there be a printing of labels!
  def go_go_go
  
    # Catch errors so the GUI doesn't just mysteriously disappear
    begin
    
      # Make sure they have a RDT username and password
      unless get_regkey_val('username') && get_regkey_val('password')
        return false unless ask_for_details( 'RDT Details:' )
      end
      
      # Let's do this thing!
      m = MechReporter.new
 
      key = MechReporter.cntxify( @input.text )
      
      labelfile = m.save_labels( [ key ], 3 )
      
      m.print_label( labelfile )
      
    # Catch and report errors
    rescue => e
      error "#{ e.to_s.scan(/(.{1,200}\s?.{1,200}+)/).first.join($/) }\n\n#{ e.backtrace.first }"
    end
    
    # Quit
    #Gtk.main_quit
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
  
end

# Allows EXE builds without showing the GUI
MechReporter.new if defined?( Ocra )

# Let there be a GUI!
GUI = LabelPrinter.new

# Go!
GUI.run
