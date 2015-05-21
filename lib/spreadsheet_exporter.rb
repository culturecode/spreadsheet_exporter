require 'spreadsheet_exporter/csv'
require 'spreadsheet_exporter/xlsx'
module SpreadsheetExporter
  begin
    Mime::Type.register "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", :xlsx
  rescue NameError
    # No Mime::Type
  end
end
