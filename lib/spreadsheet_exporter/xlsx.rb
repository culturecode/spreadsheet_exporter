require 'write_xlsx'
require_relative 'spreadsheet'

module SpreadsheetExporter
  module XLSX
    def self.from_objects(objects, options = {})
      spreadsheet = Spreadsheet.from_objects(objects, options)
      from_spreadsheet(spreadsheet)
    end

    def self.from_spreadsheet(spreadsheet, temp_file_path = 'tmp/items.xlsx')
      # Create a new Excel workbook
      workbook = WriteXLSX.new(temp_file_path)

      # Add a worksheet
      worksheet = workbook.add_worksheet

      # Add and define a format
      headerFormat = workbook.add_format # Add a format
      headerFormat.set_bold

      # Write header row
      spreadsheet.first.each_with_index do |column_name, col|
        worksheet.write(0, col, column_name, headerFormat)
      end

      spreadsheet[1..-1].each_with_index do |values, row|
        values.each_with_index do |value, col|
          worksheet.write(row + 1, col, value)
        end
      end

      # Output the file contents and delete it
      workbook.close
      file = File.open(temp_file_path)
      output = file.read
      file.close
      File.delete(temp_file_path)

      return output
    end
  end
end
