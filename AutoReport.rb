
# Class to simplify automated reporting
class AutoReport

  # Bring in the workhorse code
  require_relative 'MechReporter'
  # Simplify date manipulation
  include DateTools
  # Allow access to the Registry
  include RegistryTools
  
  # Provide access to Class contents
  attr_accessor :m, :name, :email_to, :email_bcc, :files, :body, :checkname, :subject, :automation, :test

  # Set up the Class
  def initialize( email_to=Passwords::MyName )
  
    # Reporting tool
    self.m = MechReporter.new
    
    # Report name
    self.name = File.basename( $0, '.*' )
    
    # File to check for failed reports
    self.checkname = "#{ documents_path }\\#@name.txt"
    
    # Email content
    self.email_to = email_to
    self.email_bcc = []
    self.files = []
    self.body = 'Automated Email'
    self.subject = nil
    
    # Check automation mode
    case ARGV[0]
    when /check/i
      puts 'Check mode active. Setting Automate mode: true'
      self.automation = true
      if ( File.read( checkname ) == today.to_s rescue false )
        puts 'Report already completed today'
        exit
      end
    when /automate/i
      puts 'Automated'
      self.automation = true
    when /test/i
      self.automation = true
      self.test = true
      self.email_to = Passwords::MyName
    else
      puts 'Not Automated'
      self.automation = false
    end
    
  end # initialize

  # Wrapper for the report code, automatically handles errors and automated emailing
  def run
    tried = false
  
    # Error handler
    begin
    
      # This is where the real work gets done
      yield m
      
    # Catch errors
    rescue => err
    
      # Display error in console
      puts ( error_details = "#{ err.message }\n\n#{ err.backtrace.join($/) }" )
    
      if tried
      
        # Email failure message if automated
        logerror( "#{ Time.now }\n#{ error_details }" ) if automation && !test
        exit
        
      else
      
        tried = true
        
        # Only retry if automated
        if automation
          retry
        else
          # Only hold the console window open if not automated
          gets
          exit
        end
        
      end # tried
      
    end # errorhandler
    
    if automation
      
      begin
      
        # Send email and log success if automated and successful
        sendmail
        File.write( checkname, today )
      
      # Catch failures like invalid email addresses
      rescue => err
        
        # Display error in console
        puts ( error_details = "#{ err.message }\n\n#{ err.backtrace.join($/) }" )
        
        # Email failure message
        logerror( "#{ Time.now }\n#{ error_details }" ) if automation && !test
        raise err
        
      end # errorhandler
      
    end # if automation
    
  end # run
  
  def sendmail
    m.send_email( files, subject || name, email_to, body, ( test ? [] : email_bcc ) )
  end
  
  def logerror( error )
    m.send_email( [], "Failure - #{ name } #{ date_str }", Passwords::ErrorAlert, error ) rescue nil
  end
  
end