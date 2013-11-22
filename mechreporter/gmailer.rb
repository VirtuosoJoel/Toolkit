require 'gmail'                                # Send gmails automatically (includes mail gem)

class MechReporter

  #
  # Send an email through a dedicated gmail account
  #
  # @param [String, Array<String>] attachments the filename(s) to attach
  # @param [String] subject_var the email subject
  # @param [String, Array<String>] to_ary the "To" email address(es)
  # @param [String] body_var the email body
  #
  
  def send_gmail(attachments=[], subject_var='Automated Email', to_ary=Passwords::MyName, body_var='Automated email')
    
    # Standardise email addresses into an array in case a string was passed through
    to_ary = [ to_ary ].flatten.compact
    
    # Remove whitespace and add domain if required.
    # This allows more readable and DRY names to be passed into the method.
    to_ary.map! { |name| ( name.include?('@') ? name : name + Passwords::EmailSuffix ).gsub(/\s/,'') }
    
    # Make sure attachments is an array
    attachments = [ attachments ].flatten if attachments.size != 0
    
    # Send the email
    Gmail.connect( Passwords::GmailUser, Passwords::GmailPass ) do |g|
      g.deliver do
        to to_ary
        subject subject_var
        attachments.each { |filename| add_file(filename) } unless attachments.empty?
        body body_var
      end
    end

  end
  

end