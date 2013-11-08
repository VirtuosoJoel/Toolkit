require 'date'

#
# Calculates commonly used Dates, making other code more human-readable
#

module DateTools
  
  #
  # Helper method which formats a Date as a String
  #
  # @param [Date] date the Date to format, default is today.
  # @return [String] the date formatted as 'd-m-yyyy'
  #
  
  def date_str( date=today )
    date.strftime( '%-d-%-m-%Y' )
  end
  
  #
  # The last day of the preceding month
  #
  # @return [Date]
  #
  
  def end_of_last_month
    end_of_this_month << 1
  end

  #
  # The end of month of the given Date
  #
  # @param [Date] date the date to use as a starting point
  # @return [Date]
  #
  
  def end_of_month( date )
    ( start_of_month( date ) >> 1 ) - 1
  end

  #
  # The last Date of the current month
  #
  # @return [Date]
  #
  
  def end_of_this_month
    end_of_month( today )
  end
  
  #
  # The last workday including or prior to the given date, with an optional array of working days (1 is Monday)
  #
  # @param [Date] date the date to use as a starting point
  # @param [Array<Fixnum>] working_days the array of weekdays (1-7) included in a work schedule
  # @return [Date] the first working date found going backwards from the given date (inclusive)
  #

  def last_workday( date, working_days = [1,2,3,4,5] )
    date-=1 until working_days.include?( date.cwday )
    date
  end
  
  #
  # The next workday after the given date, with an optional array of working days (1 is Monday)
  #
  # @param [Date] date the date to use as a starting point
  # @param [Array<Fixnum>] working_days the array of weekdays (1-7) included in a work schedule
  # @return [Date] the next working date after the given date.
  #

  def next_workday( date, working_days = [1,2,3,4,5] )
    loop do
      date+=1
      break if working_days.include?( date.cwday )
    end
    date
  end
  
  #
  # The first day of the preceding month
  #
  # @return [Date]
  #
  
  def start_of_last_month
    start_of_this_month << 1
  end
  
  #
  # The start of month of the given Date
  #
  # @param [Date] date the date to use as a starting point
  # @return [Date]
  #

  def start_of_month( date )
    date - date.day + 1
  end
  
  #
  # The first Date of the current month
  #
  # @return [Date]
  #
  
  def start_of_this_month
    start_of_month( today )
  end

  #
  # Convert a string into a Date object
  #
  # @param [String] input the string to extract the Date from
  # @return [Date] the String as a Date
  # @return [nil] nil if the String cannot be converted to a Date
  #

  def string_to_date( input )
    Date.strptime( input.tr( '/', '-' ) ,'%Y-%m-%d' ) rescue nil
  end
  
  #
  # Today's Date
  #
  # @return [Date]
  #
  
  def today
    Date.today
  end

  #
  # Yesterday's Date
  #
  # @return [Date]
  #
  
  def yesterday
    today - 1
  end
  
end