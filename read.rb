require 'spreadsheet'
# http://spreadsheet.rubyforge.org/GUIDE_txt.html

Spreadsheet.client_encoding = 'UTF-8'
# Opens
book = Spreadsheet.open './files/example.xls'
# Accessing all worksheets in a excel file
book.worksheets
# Access specific sheet in the excel file (name or index)
sheet1 = book.worksheet('Sheet1')
sheet2 = Book.worksheet(0)

sheet1.each do |row|
  puts row
end

# Omit 2 lines, starts from line 3
sheet2.each 2 do |row|
  # do something interesting with a row
end

# access specific row
row = sheet1.row(3)
# access specific data, like an array
# returns a String, a Float, an Integer, a Formula, a Link or a Date or DateTime object - or nil if the cell is empty.
row[0]