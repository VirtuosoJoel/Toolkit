require 'win32/registry' # Registry access
require 'crypt/blowfish' # Encryption

module RegistryTools

  def get_regkey_val(keyname)
    keyval = Win32::Registry::HKEY_CURRENT_USER.open('SOFTWARE\\ReDeTrackRuby')[keyname] rescue false
    keyname == 'password' && keyval ? decrypt(keyval) : keyval
  end

 def set_regkey_val(keyname, keyval)
    keyval = encrypt keyval if keyname == 'password'
    Win32::Registry::HKEY_CURRENT_USER.create 'SOFTWARE\\ReDeTrackRuby'
    Win32::Registry::HKEY_CURRENT_USER.open('SOFTWARE\\ReDeTrackRuby', Win32::Registry::KEY_ALL_ACCESS).write(keyname, Win32::Registry::REG_SZ, keyval)
    keyname == 'password' && keyval ? decrypt(keyval) : keyval
  end
  
  def encrypt(make_it_weird)
    Crypt::Blowfish.new( Passwords::EncryptionKey ).encrypt_string make_it_weird
  end

  def decrypt(wtf_is_this)
    Crypt::Blowfish.new( Passwords::EncryptionKey ).decrypt_string(wtf_is_this)
  end

  def documents_path
    Win32::Registry::HKEY_CURRENT_USER.open( 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders' )['Personal'] rescue Dir.pwd.gsub('/','\\')
  end
  
end