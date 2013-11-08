
#
# Obsolete Array modification project
# Precursor to RubyExcel
#

class Excel_Sheet<Array

  def initialize( val=[] )
    val = %w(A1 B1 C1 A2 B2 C2 A3 B3 C3).each_slice(3).to_a if val == 'test'
    if !val.nil? && val[0].kind_of?( Array ) && !val.empty?
      val = val.reject { |ar| ar.count(nil) == ar.length }
      val = val.map { |ar| ar.map { |el| (el.nil? || el == '') ? nil : el.to_s.strip } }
    end
    super ( val )
  end
  
  def columns
    ensure_shape
    self[0].length
  end

  def rows
    length
  end

  def ensure_shape
    max_size = self.max_by(&:length).length
    map! { |ar| ar.length == max_size ? ar : ar + Array.new( max_size - ar.length, nil) }
  end

  def get_cols( headers )
    ensure_shape
    headers = [ headers ] unless headers.kind_of? Array
    Excel_Sheet.new transpose.select { |header,_| headers.index(header) }.sort_by { |header,_| headers.index(header) || headers.length }.transpose
  end

  def get_cols?
    puts 'takes an array of headers and returns only those columns.'
  end
  
  def get_cols!( headers )
    replace get_cols( headers )
  end

  def get_col( header )
    get_cols( header ).flatten
  end
  
  def get_col?
    puts 'takes a header and returns that column as a flattened array'
  end

  def skip_headers
    Excel_Sheet.new ( block_given? ? ( [ self[0] ] + yield( self[1..-1] ) ) : ( self[1..-1] ) )
  end

  def get_row( header, lookup_key )
    self[ get_col( header ).index( lookup_key ) ]
  end

  def get_value( val_header, lookup_header, lookup_key )
    self[ get_col( lookup_header ).index( lookup_key ) ][ self[0].index( val_header ) ] rescue nil
  end

  def get_value?
    puts 'self.get_value "Value Column", "Search Column", "Find"'
  end

  def find( val )
    each { |ar| return [ self[0][ ar.index( val ) ], ar.index( val ) ] if ar.include? val }
    nil
  end

  def find?
    puts 'Returns [ "Header", row ] or nil.'
  end

  def filter( header, regex, switch=true )
    fail ArgumentError, "#{regex} is not valid Regexp" unless regex.class == Regexp
    idx = self[0].index header
    fail ArgumentError,  "#{header} is not a valid header" if idx.nil?
    operator = ( switch ? :=~ : :!~ )
    Excel_Sheet.new skip_headers { |xl| xl.select { |ar| ar[idx].send( operator, regex ) } }
  end
  
  def filter!(*p)
    replace filter(*p)
  end
  
  def filter?
    puts 'self.filter "Header", /Regex/'
  end
  
  def unique( header )
    idx = self[0].index header
    fail ArgumentError,  "#{header} is not a valid header" if idx.nil?
    Excel_Sheet.new skip_headers { |xl| xl.uniq { |ar| ar[idx] } }
  end
  
  def unique! *p 
    replace unique *p
  end
    
  def to_s
    map { |ar| ar.map { |el| "#{el}".strip.gsub( /\s/, ' ' ) }.join "\t" }.join $/
  end

  def to_s!
    replace to_s
  end

  def upcase
    map { |row| row.map{ |cell| cell.nil? ? cell : "#{cell}".upcase } }
  end

  def upcase!
    replace upcase
  end
  
  def strip
    map { |ar| ar.map { |el| el.nil? ? el : el.to_s.strip } }
  end

  def strip!
    replace strip
  end
  
  def is_float?( val )
    !!Float(val) rescue false
  end
  
end
