require_relative '../MechReporter'
require_relative '../WinBoxes'

#
# Quick hack to get rapid job unalloction
#

class UnallocateTool
  include WinBoxes
  
  UnallocateLinkBase = 'http://centrex.redetrack.com/redetrack/bin/unassignjob.php?si=222&dep=Centrex+Computing+Services&bc=&seq=1&confirmonly=y&store=y&create_box_row_on_route=y'
  
  attr_accessor :query, :m
  
  def initialize
    
    # Create the query object
    self.query = RDTQuery.new UnallocateLinkBase
    
    # Create the object which links into RDT
    self.m = MechReporter.new.login
    
  end
  
  def request_loop
    
    # process all requests
    loop do
      
      # get user input
      result = inputbox 'Scan the CNTX to unallocate:', 'Unallocate'
      
      # terminate on cancel
      break if result.to_s.empty?
      
      # prepare the uri
      query[ 'bc' ] = MechReporter.cntxify( result )
      
      # send the request to the server
      m.agent.get query
      
    end # loop
    
  rescue => error
  
    errbox error.to_s
    
  end # request_loop
  
end

UnallocateTool.new.request_loop