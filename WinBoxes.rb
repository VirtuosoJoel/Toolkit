require 'win32ole'
require 'dl'

module WinBoxes

  # button constants
  VBOKONLY = 0
  VBOKCANCEL = 1
  VBABORTRETRYIGNORE = 2
  VBYESNO = 4

  # style constants
  VBCRITICAL = 16
  VBQUESTION = 32
  VBEXCLAMATION = 48
  VBINFORMATION = 64
  VBSYSTEMMODAL = 4096

  # return code constants
  VBOK = 1
  VBCANCEL = 2
  VBABORT = 3
  VBRETRY = 4
  VBIGNORE = 5
  VBYES = 6
  VBNO = 7
  
  def inputbox(message, title = '', default = '')
    WIN32OLE.new('ScriptControl').tap { |sc| sc.language = 'VBScript' }.eval(%Q|Inputbox("#{ message.gsub( /\n/, '" & vbCrLf & "' ) }", "#{ title }", "#{ default }")|)
  end

  def popup( message, msgtitle="FYI", delay=1 )
    WIN32OLE.new('WScript.Shell').popup(message, delay, msgtitle)
  end

  def msgbox( txt, title = 'Message', buttons = VBOKONLY )
    DL::CFunc.new(DL.dlopen('user32')['MessageBoxA'], DL::TYPE_LONG, 'MessageBox').call([0, txt, title, buttons].pack('L!ppL!').unpack('L!*'))
  end

  def errbox( txt )
    puts txt
    msgbox txt, 'Error', VBCRITICAL + VBSYSTEMMODAL
  end

end