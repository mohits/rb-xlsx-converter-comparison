# This script uses Apache POI for creating an XLSX file with 1 million rows, 20 cells each
#  so that we can use it for comparison with other options

t1 = Time.now
puts t1

require 'xlsxtream'

Xlsxtream::Workbook.open('xstream_1m-rows.xlsx') do |xlsx|
  xlsx.write_worksheet('Results') do |sheet|
    # Add a million rows to the sheet
    row = Array.new(20)
    (0...1_000_000).each { |_r|
      (0..19).each { |x| # Add 20 columns to the row
        cell = (rand * 1000).to_i # store a 4 digit random value
        row[x] = cell
      }
      sheet << row
    }
  end
end

t2 = Time.now
puts t2

puts "Duration: #{t2 - t1}."
puts 'Press enter to finish.' # wait so that we can read off the memory from Task Manager or such, if we want
gets
