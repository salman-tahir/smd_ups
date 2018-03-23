# 
# cls_net.rb
# 
# Created on Sep 17, 2007, 3:50:17 PM
# 
# To change this template, choose Tools | Templates
# and open the template in the editor.
 
require "net/ping"
   include Net

class NetPing
  
  attr_accessor :host, :webpage 
  
  PingTCP.econnrefused = true
  
  def initialize(strHost, webpage)
    @host = strHost
    @webpage = webpage
  end
  
  def check_host()

    begin
      pe = PingExternal.new(@host)
      if pe.ping
       return 1
      else
       return pe.exception
      end

    rescue Exception =>ex
      puts ex
    end

  end
    
  def ping_host()
  
    begin
      ph = PingHTTP.new(@webpage)
      if ph.ping?
        return 1
      else
        return ph.exception
      end
      
    rescue Exception => ex
      puts ex
    end
  
  end
  
end

#HOST = 'secure.goldstandard.com'
#LOGON_PAGE = 'https://secure.goldstandard.com/cp-login.asp?referrer=1&code=0&entry='
#my_net = NetPing.new(HOST,LOGON_PAGE)
#puts my_net.check_host
#puts my_net.ping_host

