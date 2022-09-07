require 'csv'
require_relative 'spreadsheet'

module SpreadsheetExporter
  module CSV
    BOM = "\377\376".force_encoding("utf-16le") # Byte Order Mark so Excel displays characters correctly

    def self.from_objects(objects, humanize_headers_class: nil, **options)
      spreadsheet = Spreadsheet.from_objects(objects, humanize_headers_class: humanize_headers_class, **options).compact
      from_spreadsheet(spreadsheet)
    end

    def self.from_spreadsheet(spreadsheet, encoding: 'UTF-8', col_sep: "\t", **options)
      output = ::CSV.generate(encoding: encoding, col_sep: col_sep, **options) do |csv|
        spreadsheet.each do |row|
          csv << row
        end
      end

      BOM + output.encode!('utf-16le')
    end
  end
end
