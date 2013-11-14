require 'httparty'
require 'icalendar'

#
# Rolling Bank Holiday list for England & Wales
#

class BankHolidays
  attr_reader :all
  
  def initialize
    @all = Icalendar.parse( HTTParty.get( 'https://www.gov.uk/bank-holidays/england-and-wales.ics' ) ).first.events.map &:dtstart
  end
  
  def next_working_day( date, weekdays=[1,2,3,4,5] )
    loop do
      date +=1
      break if weekdays.include?( date.cwday ) && !all.include?( date )
    end 
    date
  end
  
  def working_minutes_between( time1, time2 )
  
    # Make sure we're working with Times.
    [ time1, time2 ].each { |time| Time === time or fail ArgumentError, "Must be a Time: #{ time }" }
  
    # Put the times in the right order
    time1, time2 = [ time1, time2 ].sort
    
    # Instantiate the difference
    diff = 0
    
    # Loop until the calculation is complete
    until time1 >= time2
    
      # If it's a date we don't count
      if time1.saturday? || time1.sunday? || all.include?( time1.to_date )
        
        # Advance until the end of the day
        time1 = ( time1.to_date + 1).to_time
      
      # If it's a date we do count
      else time2 - time1 >= 86400
        
        # Calculate the next datetime in the sequence
        newdate = [ ( time1.to_date + 1 ).to_time, time2 ].min
        
        # Count the difference
        diff += ( ( newdate - time1 ) / 60 ).floor
        
        # Advance to the new datetime
        time1 = newdate
      
      end
      
    end
    
    # Return the result
    diff
    
  end
  
end