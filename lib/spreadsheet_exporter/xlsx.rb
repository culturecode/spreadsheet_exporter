require 'write_xlsx'
require_relative 'spreadsheet'

module SpreadsheetExporter
  module XLSX
    def self.from_objects(objects, options = {})
      spreadsheet = Spreadsheet.from_objects(objects, options).compact
      from_spreadsheet(spreadsheet)
    end

    def self.from_spreadsheet(spreadsheet)
      io = StringIO.new
      # Create a new Excel workbook
      workbook = WriteXLSX.new(io)

      # Add a worksheet
      worksheet = workbook.add_worksheet

      # Add and define a format
      headerFormat = workbook.add_format # Add a format
      headerFormat.set_bold

      # Write header row
      Array(spreadsheet.first).each_with_index do |column_name, col|
        worksheet.write(0, col, column_name, headerFormat)
      end

      Array(spreadsheet[1..-1]).each_with_index do |values, row|
        Array(values).each_with_index do |value, col|
          worksheet.write(row + 1, col, value)
        end
      end

      workbook.close
      io.string

    end
  end
end
