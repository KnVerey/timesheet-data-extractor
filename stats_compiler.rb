require 'rubygems'
require 'spreadsheet'
Spreadsheet.client_encoding = 'UTF-8'

def set_format(sheet, row_num)
	format = Spreadsheet::Format.new :color => :red,
										:weight => :bold,
										:size => 11
	sheet.row(row_num).default_format = format									
end

def set_headings(sheet)
	sheet[0,0] = "Request"
	sheet.column(0).width = 8
	sheet[0,1] = "Words"
	sheet.column(1).width = 8
	sheet[0,2] = "Task"
	sheet.column(2).width = 27
	sheet[0,3] = "Hours"
	sheet.column(3).width = 8
	sheet[0,4] = "Comment"
	sheet.column(4).width = 80

	set_format(sheet, 0)
end

def add_totals(storage_sheet, rows_created)
	storage_sheet[rows_created, 0] = "TOTAL"
	words = 0
	hours = 0
	storage_sheet.each do |row|
		words+= row[1].to_i unless row[1].nil?
		hours+= row[3].to_f unless row[3].nil? 
	end

	storage_sheet[rows_created, 1] = words
	storage_sheet[rows_created, 3] = hours
	set_format(storage_sheet, rows_created)
end

def pull_data(databook, storage_sheet, requests)
	rows_created = 1 #It's the heading row

	databook.worksheets.each do |month|
		month.each do |line|

			if requests.include?(line[5]) 
				storage_sheet[rows_created, 0] = line[5] #request number
				storage_sheet[rows_created, 1] = line[7] #words
				storage_sheet[rows_created, 2] = line[8] #task
				storage_sheet[rows_created, 3] = line[9] #time
				storage_sheet[rows_created, 4] = line[14] #comment

				rows_created += 1
			end
		end
	end
	add_totals(storage_sheet, rows_created)
end

last_year = Spreadsheet.open "/Users/Katrina/Box\ Documents/Work/IRB/Last\ year.xls"
this_year = Spreadsheet.open "/Users/Katrina/Box\ Documents/Work/IRB/This\ year.xls"

requests = ["IR51844","IR51837","IR53453","IR52985","IR53940","IR51839","IR53984","IR51840","IR53894","IR52323","IR51836","IR52319","IR51841","IR52321","IR51843","IR52320","IR51842","IR52986","IR52720","IR52518","IR52318","IR51838"]

handbook_stats = Spreadsheet::Workbook.new
sheet1 = handbook_stats.create_worksheet :name => "FY 2012-2013"
sheet2 = handbook_stats.create_worksheet :name => "FY 2013-2014"

set_headings(sheet1)
pull_data(last_year, sheet1, requests)

set_headings(sheet2)
pull_data(this_year,sheet2,requests)

handbook_stats.write "/Users/Katrina/Box\ Documents/Work/IRB/handbook_stats.xls"
