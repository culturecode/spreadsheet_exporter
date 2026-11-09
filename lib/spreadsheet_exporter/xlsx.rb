require 'write_xlsx'
require 'stringio'
require_relative 'spreadsheet'

module SpreadsheetExporter
  module XLSX
    def self.from_objects(objects, options = {})
      spreadsheet = Spreadsheet.from_objects(objects, options).compact
      from_spreadsheet(spreadsheet)
    end

    def self.from_spreadsheet(spreadsheet)
      # Create a new Excel workbook
      output = StringIO.new
      workbook = WriteXLSX.new(output)

      # Add a worksheet
      worksheet = workbook.add_worksheet

      # Add and define a format
      headerFormat = workbook.add_format # Add a format
      headerFormat.set_bold

      # Write header row
      Array(spreadsheet.first).each_with_index do |column_name, col|
        if column_name.is_a?(String)
          worksheet.write_string(0, col, column_name, headerFormat)
        else
          worksheet.write(0, col, column_name, headerFormat)
        end
      end

      Array(spreadsheet[1..-1]).each_with_index do |values, row|
        Array(values).each_with_index do |value, col|
          if value.is_a?(String)
            worksheet.write_string(row + 1, col, value)
          else
            worksheet.write(row + 1, col, value)
          end
        end
      end

      workbook.close

      return output.string
    end
  end
end
