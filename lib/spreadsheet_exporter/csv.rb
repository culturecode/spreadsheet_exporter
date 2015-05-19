require 'csv'
require_relative 'spreadsheet'

module SpreadsheetExporter
  module CSV
    BOM = "\377\376".force_encoding("utf-16le") # Byte Order Mark so Excel displays characters correctly

    def self.from_objects(objects, options = {})
      spreadsheet = Spreadsheet.from_objects(objects, options)
      from_spreadsheet(spreadsheet)
    end

    def self.from_spreadsheet(spreadsheet, temp_file_path = 'tmp/items.xlsx')
      output = ::CSV.generate(:encoding => 'UTF-8', :col_sep => "\t") do |csv|
        spreadsheet.each do |row|
          csv << row
        end
      end

      return BOM + output.encode!('utf-16le')
    end
  end
end
