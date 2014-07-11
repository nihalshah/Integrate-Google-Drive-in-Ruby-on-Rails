class Googledrive

require 'digest/md5'
require 'net/smtp'
require 'mail'
require 'net/imap'
require 'gmail'
require "rubygems"

def self.googledriveappend(username, password, spreadsheetid, text, after_row)

  if !username.nil? && !password.nil? && text.length >0


#################################################
# Logs the user into Google Drive
#################################################
      @gusername = username
      @gpassword = password

      @session = GoogleDrive.login(@gusername, @gpassword)
      @ws = @session.spreadsheet_by_key(spreadsheetid).worksheets[0]


#################################################
# cur_max_rows =  the last row number
# max_col_len = last column number
# target_rows = row to shift the data to
# between_rows = number of rows to cut-copy-paste
#################################################

      @cur_max_rows = @ws.num_rows
      @max_col_len = @ws.num_cols
      @target_rows = @cur_max_rows + text.length
      @between_rows = @cur_max_rows - after_row


#################################################
# Code to cut-copy-paste the appropriate rows
#################################################


      col = 1
      @between_rows.times do
      	@max_col_len.times do
      		@ws[@target_rows,col] = @ws[@cur_max_rows,col]
      		col+=1
  		end
  		@target_rows -= 1
  		@cur_max_rows -= 1
  		col = 1
	  end


######################################################
# Makes the rows, where new data will be put, blank
# Then, appends the new data in the appropriate place
######################################################
	  @wsrow = after_row+1
     cur_row=0
     cur_col=1
     	text.length.times do
        @textcol = text[cur_row].length
        @max_col_len.times do
        	@ws[@wsrow+cur_row,cur_col] = ""
        	cur_col+=1
        end
        cur_col=1
        @textcol.times do
          @ws[@wsrow + cur_row , cur_col] = text[cur_row][cur_col-1]
          cur_col+=1
        end
        cur_row+=1
        cur_col=1
      end
      @ws.save()
     
end

end
end
