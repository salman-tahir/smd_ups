# This script instantiates the UPS class 
# Author:: Luis Lebron - now Tom Zurn 6/30/11
# Copyright:: (c) 2007 F1Networks

str_path = File.expand_path(File.dirname(__FILE__))

require 'logger'
require str_path + '/clsUps.rb'

#this section is for testing if sw can make connection with web site
#require str_path + '/clsNet.rb'

#str_host = 'www.ups.com'
#str_logon = 'http://www.ups.com/WebTracking/track?loc=en_US'

#Check to see if we can get to the UPS web site
#my_net = NetPing.new(str_host,str_logon)

#if my_net.ping_host().to_i != 1
#  log.fatal("Could not connect to #{str_logon}") 
#  puts "Could not connect to #{str_logon}."
#  exit
#end 

log = Logger.new('SMDUps.log', 5, 10*1024)
log.info("SMD Ups started ")

#Start the main program
my_smd_ups = SMDUPS.new()
#my_smd_ups.get_tracking_nums()
my_smd_ups.testit

puts 'SMD Ups completed'
