require 'creek'
require 'spreadsheet'
require 'pp'

#Print directions

puts ' =============================== '
puts '========== ssn scraper =========='
puts ' =============================== '
puts 'INSTRUCTIONS:'
puts '1) Place this file in the same directory as your .xls and .xlsx files'
puts '2) Run the file'
puts '3) SSN info found in excel files will be written to output.xls'


#Get a list of all the files

xls_pattern = File.join("**", "*.xls")
xls_files = Dir.glob(xls_pattern)

xlsx_pattern = File.join("**", "*.xlsx")
xlsx_files = Dir.glob(xlsx_pattern)

output_data = []



# Load and scrape the xls files

xls_files.each do |filename|
  puts "Reading #{filename}..."
  workbook = Spreadsheet.open 'data/034 Summer Cont Students Samir Transcripts1.xls'
  worksheets = workbook.worksheets

  worksheets.each do |worksheet|
    worksheet.rows.each_with_index do |row, row_index|
      row_cells = row.to_a.map{ |v| v.methods.include?(:value) ? v.value : v }
      row_cells.each_with_index do |cell, column_index|
        output_data << [cell,filename,worksheet.name,row_index,column_index]  if cell and cell.to_s.gsub('-','') =~ /^\d\d\d\d\d\d\d\d\d$/
      end
    end
  end
end



# Load and scrape the xlsx files

xlsx_files.each do |filename|
  puts "Reading #{filename}..."
  workbook = Creek::Book.new filename
  worksheets = workbook.sheets

  worksheets.each do |worksheet|
    worksheet.rows.each_with_index do |row, row_index|
      row_cells = row.values
      row_cells.each_with_index do |cell, column_index|
        output_data << [cell,filename,worksheet.name,row_index,column_index] if cell and cell.to_s.gsub('-','') =~ /^\d\d\d\d\d\d\d\d\d$/
      end
    end
  end
end


#Save the data

Spreadsheet.client_encoding = 'UTF-8'

book = Spreadsheet::Workbook.new
sheet1 = book.create_worksheet :name => 'Results'

sheet1.row(0).push 'ssn', 'file', 'sheet name','row index', 'column index'
output_data.each_with_index do |row, index|
  row.each do |value|
    sheet1.row(index+1).push value
  end
  # sheet1.row(index+1).push row
end

book.write 'output.xls'