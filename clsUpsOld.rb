# This program gets tracking numbers from the SMD Ops database and updates their shipping status
# Author:: Luis Lebron - now Tom Zurn 6/30/11
# Revised:: Jul 31, 2009
# Copyright:: (c) 2007 F1Networks

str_path = File.expand_path(File.dirname(__FILE__))

require 'win32ole'
require 'date'
require 'mechanize'
require 'hpricot'
require str_path + '/clsSQL.rb'

class SMDUPS
 attr_accessor(:doc, :file_string, :tracking_nums, :status_codes, :filename, :filepath, :savefile)

  DB_FILE = 'C:\Ops\Data\SMD Data.mdf'

  def initialize
    @status_codes = Hash["In Transit", 1, "In Transit - On Time", 1, "Delivered", 2, "Exception", 3, "Pickup", 4, "Manifest Pickup", 5, "Bad Number", 6, "Not Available", 7, "Billing Information Received", 8, "Returned to shipper", 9, "Voided Information Received", 10, "Out For Delivery", 11]
    @doc = ""
    @file_string = ""
    @filepath = "C:/Ops/Customers/"
    @savefile = 1
  end

  #Get the tracking numbers from the SMD OPS database and process them
  def get_tracking_nums()

    begin
      @tracking_nums = {}
      db = SQLDb.new(DB_FILE)
      db.open
      db.query("SELECT TrackingNum, SalesContract.ID, Name, SSNUM from SalesContract, Patient WHERE ShippingSvc = 1 AND NOT TrackingNum IS NULL AND ((ShippingSts <> 2 AND ShippingSts <> 9) OR ShippingSts IS NULL) AND SalesContract.LocatorID = Patient.ID")

      rows = db.data

      rows.each do |row|
        if row[3].to_s.rstrip.length == 9 #Check for the length of the SSN
          @tracking_nums[row[0].strip().gsub(' ','').upcase]= row[1]
          process_tracking_nums(row[0], row[3], row[2])

        else
          log_error("Incorrect SSN for " + row[2].to_s)
        end
      end
      db.close
      
      save_results()
      

    rescue Exception =>ex
      puts ex.to_s
      log_error(ex.to_s)
    end

  end

   def testit
    agent = Mechanize.new
      page = agent.get('http://wwwapps.ups.com/WebTracking/track')
      #ups_form = page.forms[0]
      #puts "0: " + ups_form.name
      ups_form = page.forms[1]
      puts "1: " + ups_form.name
      puts ups_form.trackNums
      #ups_form = page.forms[2]
      #puts "2: " + ups_form.name
      #ups_form = page.forms[3]
      #puts "3: " + ups_form.name
  end
  
  private
  #Submit the tracking numbers to the UPS website and get the results page
  def process_tracking_nums(trackingNums, str_ssn, str_name)
          
    begin
      agent = Mechanize.new
      page = agent.get('http://wwwapps.ups.com/WebTracking/track')
      ups_form = page.forms[1]
            
      #if formnumber changes on website, sw crashes on next line*******
      ups_form.trackNums = trackingNums
      #****************************************************************
      
      #ups_form.checkboxes.name('AgreeToTermsAndConditions').check
      page = agent.submit(ups_form, ups_form.buttons[0])
      @doc = Hpricot(page.body)
      parse_results(trackingNums.rstrip)

      if @savefile == 1
        save_file(str_ssn, str_name)
      end

    rescue Exception => ex
      puts ex.to_s
      log_error(ex.to_s)
    end

  end

  #Parse the results page to get the shipping status
  def parse_results(trackingNums)
    
    @savefile = 0

    if @doc.search("p[@class=error]").length > 0
     if @doc.search("p[@class=error]").first.inner_text.strip().include?('UPS could not locate')
        @file_string <<trackingNums
        @file_string << ",Bad Number" + "\n"
        @savefile = 0
        return
      end
    end
    
    if @doc.search("table[@class=dataTable]").length > 0
      @file_string << trackingNums + ", "
      #puts @doc
      table =@doc.search("table[@class=dataTable]")
      (table/"a").remove
      #puts table

      first_row = table.search("tr:eq(1)")
      status = first_row.search("td:eq(3)").inner_text.gsub('&nbsp;',"").gsub("\n","")
      #puts "Status: " + status
      @file_string << get_result_string(status) + "\n"
      @savefile = 1
    end
  end

  def get_result_string(tdstring)
    #puts "get_result " + tdstring
    @status_codes.each_key do |key|
      if tdstring.include?(key)
        return key.to_s
      end
    end
  end


  #Update the SMD OPS database.
  #Uses a vb executable to send the information to the server
  def save_results()
    #puts @file_string
    stime = Time.now.strftime("%m/%d/%Y %H:%M:%S")
    db = SQLDb.new(DB_FILE)
    db.open
    
    begin
      @file_string.each_line do |line|
        arrParts = line.split(",")
        str_tracking_num = arrParts[0].to_s.strip().gsub(' ','')
        strStatus = arrParts[1].to_s.strip().squeeze(" ")
        id = @tracking_nums[str_tracking_num]
        
        if @status_codes[strStatus] == nil
          str_status_code = 0
        else
          str_status_code = @status_codes[strStatus]
        end

        puts "SalesID: " + id.to_s + " Status: " + str_status_code.to_s

        #Pass the data to the vb executable
        #system("UpdateShipping.exe #{id},#{str_status_code}")
        ssql = "INSERT INTO UPS Values('" + id.to_s + "','" + stime + "'," + str_status_code.to_s + ",'false','true','false')"
        db.execute(ssql)
      end

    rescue Exception => ex
      puts ex.to_s
      log_error(ex.to_s)
    end
    
    db.close
  end

  #Save the resulting html
  def save_file(str_ssn, str_name)
    #file naming convention C:\Ops\Customer\SSN\Name_date_confirm.html

    begin
      str_filepath = @filepath + str_ssn +'/'+ str_name.rstrip + '_' + Time.now.strftime("%Y%m%d%H%M%S") + '_confirm.html'

      if !File.exist?(@filepath + str_ssn +'/')
        Dir.mkdir(@filepath + str_ssn +'/')
      end

      table =@doc.search("table[@class=dataTable]")
      (table/"a").remove
     
      a_file = File.new(str_filepath , "w")
      a_file << "<table>"
      a_file << table.inner_html().squeeze(' ')
      a_file << "</table>"
      a_file.close
      
    rescue SystemCallError
      puts "Could not create file: " + $!
      log_error("Could not create file")
    rescue Errno::ENOENT
      puts "No such file or directory"
      log_error("No such file or directory")
    rescue Errno::EACCESS
      puts "Permission denied"
      log_error("Permission denied")
    rescue Exception => ex
      puts ex.to_s
      log_error(ex.to_s)
    end

  end

  #Log any errors
  def log_error(str_error)
    if File.exist?('SMDUps.log')
      file = File.open('SMDUps.log', File::WRONLY |File::APPEND)
      log = Logger.new(file)
      log.fatal(str_error)
      log.close
    end
  end
end