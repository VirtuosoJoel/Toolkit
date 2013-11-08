require 'watir-webdriver' #Browser
require 'nokogiri' #Parse HTML with speed
require 'win32ole' #Excel
require 'dl' #Lots of handy stuff
require 'date' #Work with Dates
require_relative 'rubyexcel/lib/lib/rubyexcel' # My shiny data handling gem
require_relative 'Passwords' # Passwords & personal info
require_relative 'Passwords' # Registry & encryption

class RDT

  include RegistryTools

  def self.get_end_of_last_month
    Date.today - Date.today.day
  end

  def self.get_start_of_last_month
    Date.today - Date.today.day + 1 << 1
  end

  attr_accessor :driver, :download_directory, :current_downloads, :test, :user_pass
  
  def initialize(test = false, anonymous = false)
    ocra_build if defined?(Ocra)
    
    if anonymous
    
      self.user_pass ||= []
      
      puts 'Please enter RDT username:'
      user_pass[0] = gets.chomp
      exit if user_pass[0] == ''
      
      puts 'Please enter RDT password:'
      user_pass[1] = gets.chomp
      exit if user_pass[1] == ''
      
    else
    
      self.user_pass = [ get_regkey_val( 'username' ), get_regkey_val( 'password' ) ]
      
    end

    self.test = test
  end
  
  def start(use_profile = false, downloads = false, human_usable = false)
    if driver.nil?

      use_profile ? profile = Selenium::WebDriver::Firefox::Profile.from_name('default') : profile = Selenium::WebDriver::Firefox::Profile.new
      profile.native_events = false
      profile['permissions.default.image'] = 2 unless human_usable
      profile['app.update.auto'] = false
      
      if downloads
        self.download_directory = if downloads.is_a?( String )
          downloads
        else
          "#{ RubyExcel.documents_path }/RubyDownloads".gsub( '/', '\\' )
        end
        Dir::mkdir( download_directory ) unless File.directory?( download_directory )
        profile['browser.download.folderList'] = 2
        profile['browser.download.dir'] = download_directory
        profile['browser.helperApps.neverAsk.saveToDisk'] = 'application/pdf, application/x-pdf, application/x-download'
        profile['browser.download.manager.showWhenStarting'] = false
      end
      
      client = Selenium::WebDriver::Remote::Http::Default.new
      client.timeout = 120
      
      puts 'Opening Firefox...'
      self.driver = Watir::Browser.new :firefox, :profile => profile, :http_client => client
    else
      driver
    end
  end
  
  def wait_for_download( init = false )
    if init
      self.current_downloads = Dir.glob( ( download_directory + '\\*.pdf' ).gsub('\\','/') )
    else
      
      #Find the newest file
      difference = [1,2]
      until difference.size == 1
        difference = Dir.glob( ( download_directory + '\\*.pdf' ).gsub('\\','/') ) - current_downloads
        sleep 0.1
      end
      file_name = difference.first
      #puts "Found new file: #{file_name}"
      
      new_size = 0
      current_size = -1
      
      #Wait for file size to stop changing
      until current_size == new_size
        current_size = File.size file_name
        sleep 0.1
        new_size = File.size file_name
        sleep 0.1
      end
      
      #Extra delay so the O/S doesn't think the file is corrupt...
      sleep 1
      file_name.gsub('/','\\')
      
     end
  end
  
  def login(tool='Reports')
  
    server = ( test ? Passwords::RDTTestDomain : Passwords::RDTCoreDomain )
    start if driver.nil?
    
    driver.execute_script('window.onbeforeunload = null')
    
    driver.goto server
    
    if driver.button(:value => 'Login').present?
      [ ['org', Passwords::OrgID], ['password', user_pass[1]], ['username', user_pass[0]] ].each { |ar| driver.text_field(:id => ar[0]).when_present.set ar[1] }
      driver.button(:value => 'Login').click if driver.button(:value => 'Login').present?
      driver.button(:value => 'Logout').wait_until_present
    end
    
    driver.div(:id => 'content').links.each { |link| (driver.goto "#{ server }#{ link.attribute_value('onclick').split("'")[1] }"; break link) if link.text == tool }
    
    fail NameError, "#{ tool } not found" if driver.url == server
    
  end
    
  def custom_command(array_of_arrays)
    array_of_arrays.each { |ar| driver.send( *ar[0..2] ).send( *ar[3..-1] ) }
  end
  
  def close
    driver.close rescue nil
  end
  
  def ocra_build
    Crypt::Blowfish.new('1').encrypt_string('Moose')
    Watir::Browser.new.close
    exit
  end
  
  def create_array(inputvalue)
    var = inputvalue.gsub(/\t/,' ').split(/\n/)
    var.delete ''
    var
  end
  
  def print_label ( file_name, copies=1 )
    printer = Win32::Registry::HKEY_CURRENT_USER.open('Software\Microsoft\Windows NT\CurrentVersion\Windows')['Device'].split(',').first
    acrobat = Win32::Registry::HKEY_LOCAL_MACHINE.open('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe')['']
    cmd = %Q|"#{acrobat}" /n /s /h /t "#{ file_name.gsub('/', '\\') }" "#{ printer }"|.gsub('\\','\\\\')
    copies.times do
      sleep 0.3
      Thread.new { system cmd }
    end
  end
  alias print_labels print_label

end