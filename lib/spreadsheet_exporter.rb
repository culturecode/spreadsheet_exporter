require_relative './spreadsheet_exporter/csv'
require_relative './spreadsheet_exporter/xlsx'
require 'active_support'
require 'active_support/core_ext/object/json'

module SpreadsheetExporter
  begin
    Mime::Type.register "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", :xlsx
  rescue NameError
    # No Mime::Type
  end
end
