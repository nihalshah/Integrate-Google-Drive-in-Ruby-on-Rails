class Googledrive

require 'google_drive'
require "rubygems"


#################################################
# Logs in the user by making a connection via the
# Googledrive Gem
# Returns the sessions created.

#   ----------PARAMETERS-------------
# username and password : Self explanatory.
#
#################################################


def self.login(username, password)

  if !username.nil? && !password.nil?

      @gusername = username
      @gpassword = password

      @login = GoogleDrive.login(@gusername, @gpassword)

      return @login
  end
end



########################################################
# Inserts data into appropriate row / column
# as specified in the params.

# -------------PARAMETRS--------------

# login = login session of Googledrive
# text = "row / column"
# insert = The row or column number to insert in
# data = The text to insert in the given row(s) / column(s)
#########################################################



def self.insert(login, spreadsheetid, sheetnumber, text, insert, data)

@ws = login.spreadsheet_by_key(spreadsheetid).worksheets[sheetnumber]
   

#################################################
# cur_rows =  the last row number
# max_col_len = last column number
# target_rows = row to shift the data to
# between_rows = number of rows to cut-copy-paste
#################################################

if text == "row"

      @cur_row = @ws.num_rows
      @max_col_len = @ws.num_cols
      @target_rows = @cur_row + data.length
      @between_rows = @cur_row - insert + 1


#################################################
# Code to cut-copy-paste the appropriate rows
#################################################


      col = 1
      @between_rows.times do
      	@max_col_len.times do
      		@ws[@target_rows,col] = @ws[@cur_row,col]
      		col+=1
  		end
  		@target_rows -= 1
  		@cur_row -= 1
  		col = 1
	  end


######################################################
# Makes the rows blank - where new data will be put 
# Then, appends the new data in the appropriate place
######################################################
     cur_row = insert
     cur_col = 1
     index = 0
     data.length.times do
       @textcol = data[index].length
       @max_col_len.times do 
       @ws[cur_row, cur_col] = ""
       cur_col+=1
     end
        cur_col=1
        @textcol.times do
          @ws[cur_row , cur_col] = data[index][cur_col-1]
          cur_col+=1
        end
        cur_row += 1
        cur_col = 1
        index += 1
      end

     


elsif text == "column"


########################################################
# cur_max_rows =  the current last row number
# max_col_len = last column number
# target_col = column number where data will be shifted
# between_col = number of columns to cut-copy-paste
########################################################


  @cur_max_rows = @ws.num_rows
  @max_col_len = @ws.num_cols
  @between_col = @max_col_len - insert + 1
  @current_col = @max_col_len
  @target_col = @max_col_len + data.length


#################################################
# Code to cut-copy-paste the appropriate columns
################################################# 
  cur_row = 1
  @between_col.times do
    @cur_max_rows.times do
      @ws[cur_row, @target_col] = @ws[cur_row, @current_col]
      @ws[cur_row, @current_col] = ""
      cur_row += 1
  end
  @target_col -= 1
  @current_col -= 1
  cur_row = 1
end

  cur_row = 1

######################################################
# Makes the columns blank - where new data will be put 
# Then, appends the new data in the appropriate column
######################################################

  len = data.length
  cur_col = insert
  index = 0
  len.times do
    textlen = data[index].length
    j=0
    textlen.times do
      @ws[cur_row,cur_col] = data[index][j]
      j += 1
      cur_row += 1 
    end
    cur_col += 1
    index += 1
    cur_row = 1
  end
end
@ws.save()
end


##############################################################
# Returns a 2 dimensional array containing all positions 
# where the the data occurs (in [row,column] format),
# in the row/column as specified.
#
#               --------PARAMETERS-----------
#
# text = "row / column"
# num = The row or column number to search in
# data = The text to search for in the given row / column

# NOTE: Whatever data You want to search for, remember that
# Spreadsheets treat anything as a STRING. 
# So the DATA argument must be enclosed in double quotes.
##############################################################



def self.find(login, spreadsheetid,sheetnumber, text, num, data)
  
@ws = login.spreadsheet_by_key(spreadsheetid).worksheets[sheetnumber]

###############################
# Iterates thorugh the row and
# records all positions where
# the given data occurs
###############################
if text == "row"
  max_col = @ws.num_cols
  col=1
  ret_col = Array.new
  index = 0

  max_col.times do
    if @ws[num , col] == data
      ret_col[index] = [num,col]
      index+=1
    end
    col+=1
  end
  return ret_col


###################################
# Iterates thorugh the columns and
# records all positions where
# the given data occurs
###################################


elsif text == "column"
max_row = @ws.num_rows
row = 1
ret_row = Array.new
index = 0
max_row.times do
  if @ws[row,num] == data
    ret_row[index] = [row,num]
    index+=1
  end
  row+=1

end
return ret_row

end

end


#############################################################
# Updates a given row(s) / column(s)
#
#           --------PARAMETERS-----------
#
# text = "row / column"
# num = The row or column number to update
# data = The text to update in the given row(s) / column(s)
##############################################################

def self.update(login, spreadsheetid, sheetnumber, text,  num, data)

  @ws = login.spreadsheet_by_key(spreadsheetid).worksheets[sheetnumber]

#########################################################################
# Checks to see if data is for only a single update or multiple updates
#########################################################################
  type = data[0].class

  if text == "row"

##############################################################
# If it's a single row update, data[0].class is NOT an Array
##############################################################
    if type != Array
      max_col = data.length
      cur_row = num
      cur_col = 1
      index = 0

    @ws.num_cols.times do 
      @ws[cur_row , cur_col] = ""
      cur_col += 1
    end
    cur_col = 1
#############################################################
# Iterates through the given data and updates the row
#############################################################
      cur_row = num
      max_col.times do
        @ws[cur_row , cur_col] = data[index]
        index += 1
        cur_col+=1
    end
##############################################################
# If it's a multi row update, data[0].class IS an Array
##############################################################
    else
      update_row = data.length
      cur_row = num 
      cur_col = 1
      outer_index = 0
##########################################################################
# First gathers the length of the 2 dimensional array,
# which gives us the information about the number of rows
# to update.
# Erases the content of all relevent rows.
# Then for each array within the 2-D array,
# the code iterates through the text and updates the rows
#
#                  ---------VARIABLES----------
#
# outer_index = Length of 2-D array / num of arrays within the 2-D array
# inner_index = The number of elements in each array of the 2-D array.
##########################################################################
      update_row.times do 
        @ws.num_cols.times do 
          @ws[cur_row , cur_col] = ""
          cur_col += 1
        end
        cur_col = 1
        innder_index = 0
        data[outer_index].length.times do 
          @ws[cur_row , cur_col] = data[outer_index][innder_index]
          innder_index += 1
          cur_col += 1
        end
        cur_row += 1
        cur_col = 1
        outer_index += 1
      end
    end



  elsif text == "column"

##############################################################
# If it's a single row update, data[0].class is NOT an Array
##############################################################


    if type != Array
      max_row = data.length
      cur_col = num
      cur_row = 1
      index = 0

      @ws.num_rows.times do 
      @ws[cur_row , cur_col] = ""
      cur_row += 1
    end

#############################################################
# Iterates through the given data and updates the row
#############################################################
      cur_row = 1
      max_row.times do 
        @ws[cur_row , cur_col] = data[index]
        index += 1
        cur_row+=1
      end
##############################################################
# If it's a multi row update, data[0].class IS an Array
##############################################################
    else
      update_col = data.length
      cur_row = 1
      cur_col = num
      outer_index = 0
##########################################################################
# First gathers the length of the 2 dimensional array,
# which gives us the information about the number of columns
# to update.
# Erases the content of all relevent columns.
# Then for each array within the 2-D array,
# the code iterates through the text and updates the columns
#
#                  ---------VARIABLES----------
#
# outer_index = Length of 2-D array / num of arrays within the 2-D array
# inner_index = The number of elements in each array of the 2-D array.
##########################################################################
      update_col.times do
        @ws.num_rows.times do 
          @ws[cur_row , cur_col] = ""
          cur_row += 1
        end
        cur_row = 1
        innder_index = 0
        data[outer_index].length.times do
          @ws[cur_row , cur_col] = data[outer_index][innder_index]
          innder_index += 1
          cur_row += 1
        end
        cur_col += 1
        cur_row = 1
        outer_index += 1
      end
    end
  end
  @ws.save()
end





#############################################################
# Deletes a given row(s) / column(s)
#
#           --------PARAMETERS-----------
#
# text = "row / column"
# num = The row or column number to delete
# after = The total number of rows / columns to delete 
#         including the current (num) row / column
##############################################################

def self.delete(login, spreadsheetid, sheetnumber, text, num, after)

  @ws = login.spreadsheet_by_key(spreadsheetid).worksheets[sheetnumber]

  if text == "row"
    max_col = @ws.num_cols
    cur_col = 1
    cur_row = num


#####################################
# Makes the specified rows blank
#####################################


    after.times do 
      max_col.times do 
        @ws[ cur_row , cur_col] = ""
        cur_col += 1
      end
      cur_row += 1
      cur_col = 1
    end
##############################################################
# Shifts the blank rows to the end of the Spreadsheet
##############################################################

    num_copy = @ws.num_rows - num 
    copy_row = num + after 
    cur_row = num
    cur_col = 1

    num_copy.times do 
      @ws.num_cols.times do 
        @ws[ cur_row , cur_col] = @ws[copy_row , cur_col]
        cur_col += 1
      end
      copy_row += 1
      cur_row += 1
      cur_col = 1
    end
##############################################################
# Takes care of Duplicate rows after the whole delete process
# " num - 1 != @ws.num_rows " takes care of the case when
# the last row is deleted.
##############################################################

    cur_row = @ws.num_rows 
    cur_col=1
    if cur_row > 1 && num - 1 != @ws.num_rows
    @ws.num_cols.times do 
      @ws[cur_row,cur_col] = ""
      cur_col +=1
    end
  end


  elsif text == "column"



#####################################
# Makes the specified columns blank
#####################################
    max_row = @ws.num_rows
    cur_col = num
    cur_row = 1
    after.times do 
      max_row.times do 
        @ws[ cur_row , cur_col] = ""
        cur_row += 1
      end
      cur_row = 1
      cur_col += 1
    end
##############################################################
# Shifts the blank columns to the end of the Spreadsheet
##############################################################

    num_copy = @ws.num_cols - num
    copy_col = num + after
    cur_col = num 
    cur_row = 1

    num_copy.times do 
      @ws.num_rows.times do
        @ws[ cur_row , cur_col] = @ws[ cur_row , copy_col]
        cur_row += 1
      end
      copy_col += 1
      cur_col += 1
      cur_row = 1
    end

    cur_col = @ws.num_cols
    cur_row = 1
##############################################################
# Takes care of Duplicate columns after the  delete process
# as well as the case when the last column is deleted.
##############################################################

    if cur_col > 1 && num - 1 != @ws.num_cols
    @ws.num_rows.times do 
      @ws[cur_row , cur_col] = ""
      cur_row += 1
    end
    end
  end

  @ws.save()
end




end
