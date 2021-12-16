# This script uses Apache POI for creating an XLSX file with 1 million rows, 20 cells each
#  so that we can use it for comparison with other options

require 'java'

t1 = Time.now
puts t1

# This assumes that the binary JARs are available as in this repository
# this ruby file
# ├───log4j
# └───poi-bin-5.1.0
#     ├───auxiliary
#     ├───lib
#     └───ooxml-lib
require_relative 'log4j/log4j-core-2.16.0.jar'
require_relative 'poi-bin-5.1.0/poi-5.1.0.jar'
require_relative 'poi-bin-5.1.0/lib/commons-codec-1.15.jar'
require_relative 'poi-bin-5.1.0/lib/commons-collections4-4.4.jar'
require_relative 'poi-bin-5.1.0/lib/commons-io-2.11.0.jar'
require_relative 'poi-bin-5.1.0/lib/commons-math3-3.6.1.jar'
require_relative 'poi-bin-5.1.0/lib/log4j-api-2.16.0.jar'
require_relative 'poi-bin-5.1.0/lib/SparseBitSet-1.2.jar'
require_relative 'poi-bin-5.1.0/poi-ooxml-5.1.0.jar'
require_relative 'poi-bin-5.1.0/poi-ooxml-full-5.1.0.jar'
require_relative 'poi-bin-5.1.0/ooxml-lib/commons-compress-1.21.jar'
require_relative 'poi-bin-5.1.0/ooxml-lib/commons-logging-1.2.jar'
require_relative 'poi-bin-5.1.0/ooxml-lib/xmlbeans-5.0.2.jar'
require_relative 'poi-bin-5.1.0/ooxml-lib/curvesapi-1.06.jar'

# Point to the Java class for the streaming file
StWb = org.apache.poi.xssf.streaming.SXSSFWorkbook

# Create new file keeping only the latest 500 records in memory
wb = StWb.new(500) 

# Add a sheet to the new workbook (must provide sheet name)
sheet = wb.createSheet('Results')

# Add a million rows to the sheet
(0...1_000_000).each {|r|
  row = sheet.createRow(r)
  (0..19).each {|x|  # Add 20 colums to the sheet
    row.createCell(x).cell_value = (rand * 1000).to_i # store a 4 digit random value
  }
}

# Finalise the file - needs a Java IO FileOutputStream to write it out
java_import java.io.FileOutputStream
fos = FileOutputStream.new('poi_1m-rows.xlsx')

wb.write(fos) # Combines data from the temporary file into the final file
fos.close # Closes the file
wb.dispose # Removes the temporary file(s) that were created along the way
 
t2 = Time.now
puts t2

puts "Duration: #{t2 - t1}."
puts 'Press enter to finish.' # wait so that we can read off the memory from Task Manager or such, if we want
gets
