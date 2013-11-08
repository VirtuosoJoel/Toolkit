puts 'Do you want to create Start Menu shortcuts? [Yn]'
response = gets.chomp
response = !!(response =~ /\AY\z/i || response == '' )

# Build a command-line instruction and .iss file for each script
( Dir[ '*.rb' ] - [File.basename($0)] ).each do |filename|

  # Determine names
  basename = File.basename filename, '.*'
  issname = "#{ basename.scan( /\b[a-z]/i ).join.upcase }.iss"
  ocraname = issname.sub(/\.iss\z/, '.txt')
  
  puts "Building #{ issname } file for #{ basename }."

  # Create Ocra command in a text file
  File.write( ocraname, %Q|ocra "#{ filename }" --output "#{ basename }.exe" --chdir-first --no-lzma --innosetup "#{ issname }"| )
  
  # Create .iss file
  File.write( issname, %Q|[Setup]\nAppName=#{ basename }\nAppVersion=0.1\nDefaultDirName={pf}\\#{ basename }#{ response ? %Q|\nDefaultGroupName=#{ basename }| : '' }\nOutputBaseFilename=#{ basename }#{ response ? %Q|\n\n[Icons]\nName: "{group}\\#{ basename }"; Filename: "{app}\\#{ basename }.exe"\nName: "{group}\\Uninstall #{ basename }"; Filename: "{uninstallexe}"| : '' }| )
  
end