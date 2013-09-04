require 'rubygems'
require 'hpricot'
require 'mechanize'
require 'open-uri'
require 'win32ole'
#require "iconv"
require "socket"
require 'cgi'
WIN32OLE.codepage=WIN32OLE::CP_UTF8


def main()
  #建立資料庫連線
	connection = WIN32OLE.new('ADODB.Connection')
  #連線開啟，開啟本地端的資料庫
  connection.Open('Provider=Microsoft.Jet.OLEDB.4.0;
                    Data Source=H:\temp\SSCI.mdb')

  #撈出待比對的期刊名清單
	rs = WIN32OLE.new('ADODB.Recordset')
	sql=%Q{SELECT * FROM original_journal_list where checkif="Y"}
	rs.Open(sql, connection)
	data = rs.GetRows.transpose
  
	#針對每筆資料進行處理
	data.each{ |item|
		strISSN=""
    strISSN2=""
    strISBN=[]
    #strISBN2=""
    #strISBN3=""
    #strISBN4=""
    strJournalFullName=""
    strJournalType=""
    
    puts "====================================================================="
		print "#{item[0]}\n#{item[1]}\n"
    item[6].split("\n").each{ |line|    
        
        #puts line
        #前三欄是數字才有可能是要的，其它忽視
        if line[0,3].to_i > 0            
            case line[0,3]
            #檢查是否有ISBN欄
            when "020"
            strISBN = strISBN + ["#{line.split(" ").at(1).split("|").at(0)}"]
            print "---111-------#{line.split(" ").at(1).split("|").at(0)}-----------\n"
            #sleep(1)
             
            #檢查是否有ISSN欄
            when "022"
                strISSN=""
                strISSN2=""
                #print "#{line[7,1]}---111--------#{line[8,1]}----------\n"
                line.split("|").each{ |column|
                   print "---222-------#{column}-----------\n"
                   if column[0,1] == "l" 
                      print "---222-------#{column[1,9]}-----------\n"
                      strISSN2=column[1,9]
                   elsif column[0,3] == "022"
                      print "---222-------#{column[7,9]}-----------\n"
                      strISSN=column[7,9]   
                   end 
                }
                if strISSN2 == strISSN
                    strISSN2 = ""
                end
                #if line[7,1] == "|" 
                #    if line[8,1] == "y"
                #        strISSN=line[9,9]
                #    elsif  line[8,1] == "l"
                #        strEISSN=line[9,9]
                #    else
                #        strISSN=line[8,9]
                #    end
                #else
                #    strISSN=line[7,9]
                #end
            #回傳期刊全名
            when "245"
                strJournalFullName = line.split("|").at(0)[7,line.split("|").at(0).length]
            #回傳期刊類別
            when "650" 
                line.split("|").each{ |column|
                   #print "---222-------#{column[0,1]}-----------\n"
                   if column[0,1] == "v"# or column[0,1] == "x"
                      strJournalType = column[1,column.length]
                   end 
                }
            end

        end        
    }
    puts strISSN ,strISSN2 ,strJournalFullName, strJournalType
    strISBN = strISBN + ["","","",""]
    if strISSN != "" or strISSN2 != "" or strISBN[0] != ""              
        sql=%Q{UPDATE original_journal_list SET ISSN="#{strISSN}",ISSN2="#{strISSN2}",ISBN="#{strISBN[0]}",ISBN2="#{strISBN[1]}",ISBN3="#{strISBN[2]}",ISBN4="#{strISBN[3]}",journal_real_name="#{strJournalFullName}",journal_type="#{strJournalType}" WHERE index=#{item[0]}}
        connection.Execute sql
    end
    #sleep(1)
  if item[0].to_i == 1228
      #sleep(10)
  end
  }
	rs.close
	connection.close
end



begin
	main
rescue => ex
	puts "\n*********************************************************************"
	puts "* get error at: #{Time.now}"
	puts "* error class: #{ex.class}"
	puts "* error class: #{ex.message}"
	puts "* error backtrace: #{ex.backtrace}"
	puts "***********************************************************************"
	num=rand(10)+5
	num.downto(1) { |i| ; print "#{i}   ";sleep(1) }
	puts "GO!"
	puts "======================================================================="
	sleep(1)
	retry
end


#language-id:eng format:book subjects:"science fiction" author:"john"