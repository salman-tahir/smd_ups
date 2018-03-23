# To change this template, choose Tools | Templates
# and open the template in the editor.

str_path = File.expand_path(File.dirname(__FILE__))
require str_path + '/clsSQL.rb'
require 'win32ole'

begin
  DB_FILE = 'C:\Ops\Data\SMD Data.mdf'
  stime = Time.now.strftime("%m/%d/%Y %H:%M:%S")
  ssql = "INSERT INTO UPS Values('333','" + stime + "',0,'false','true','false')"
  db = SQLDb.new(DB_FILE)
  db.open
  db.execute(ssql)
  db.close
#puts ssql
rescue Exception =>ex
  puts ex.to_s
end
    
puts "Hello World"
